VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmSIL 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de protocolos SIL"
   ClientHeight    =   3900
   ClientLeft      =   7845
   ClientTop       =   3180
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1095
      Left            =   30
      TabIndex        =   4
      Top             =   2160
      Width           =   4515
      Begin VB.TextBox txtProtocolo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   6
         Top             =   180
         Width           =   1935
      End
      Begin esMaskEdit.esMaskedEdit mskDataEmissao 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         MouseIcon       =   "frmSIL.frx":0000
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
      Begin esMaskEdit.esMaskedEdit mskDataValidade 
         Height          =   285
         Left            =   3270
         TabIndex        =   10
         Top             =   600
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         MouseIcon       =   "frmSIL.frx":001C
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
      Begin VB.Label lblSid 
         Height          =   225
         Left            =   3570
         TabIndex        =   13
         Top             =   210
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Validade.:"
         Height          =   225
         Index           =   3
         Left            =   2280
         TabIndex        =   9
         Top             =   660
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Emissão.:"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Protocolo..:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2055
      Left            =   30
      TabIndex        =   3
      Top             =   90
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   3625
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Protocolo"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt.Emissão"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Validade"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "sid"
         Object.Width           =   0
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   3510
      TabIndex        =   0
      ToolTipText     =   "Excluir Registro"
      Top             =   3450
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
      MICON           =   "frmSIL.frx":0038
      PICN            =   "frmSIL.frx":0054
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
      Left            =   2430
      TabIndex        =   1
      ToolTipText     =   "Editar Registro"
      Top             =   3450
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
      MICON           =   "frmSIL.frx":00F6
      PICN            =   "frmSIL.frx":0112
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
      Left            =   1350
      TabIndex        =   2
      ToolTipText     =   "Novo Registro"
      Top             =   3450
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
      MICON           =   "frmSIL.frx":026C
      PICN            =   "frmSIL.frx":0288
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
      Left            =   2430
      TabIndex        =   11
      ToolTipText     =   "Gravar o Registro"
      Top             =   3450
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmSIL.frx":03E2
      PICN            =   "frmSIL.frx":03FE
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
      Left            =   3510
      TabIndex        =   12
      ToolTipText     =   "Cancelar Edição"
      Top             =   3450
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmSIL.frx":07A3
      PICN            =   "frmSIL.frx":07BF
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
Attribute VB_Name = "frmSIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodReduz As Long, Evento As String

Private Sub cmdAlterar_Click()
If txtProtocolo.Text = "" Then
    MsgBox "Nada a alterar.", vbCritical, "Erro"
    Exit Sub
End If

Evento = "Alterar"
Eventos "INCLUIR"

End Sub

Private Sub cmdCancel_Click()
Evento = ""
Eventos "INICIAR"
If lvMain.ListItems.Count > 0 Then
    lvMain.ListItems(1).Selected = True
    Le
End If

End Sub

Private Sub cmdExcluir_Click()
If txtProtocolo.Text = "" Then
    MsgBox "Nada a excluir.", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Excluir este protocolo?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "delete from sil where sid=" & Val(lblSid.Caption)
    cn.Execute Sql, rdExecDirect
    Limpa
    CarregaLista
    
    Evento = ""
    Eventos "INICIAR"
    If lvMain.ListItems.Count > 0 Then
        lvMain.ListItems(1).Selected = True
        Le
    End If
End If

End Sub

Private Sub cmdGravar_Click()
Dim nCodReduz As Long, nSid As Long
If txtProtocolo.Text = "" Then
   MsgBox "Digite o nº do protocolo.", vbCritical, "Atenção"
   Exit Sub
End If

If Not IsDate(mskDataEmissao.Text) Then
   MsgBox "Data de emissão inválida.", vbCritical, "Atenção"
   Exit Sub
End If

If Not IsDate(mskDataValidade.Text) Then
   MsgBox "Data de validade inválida.", vbCritical, "Atenção"
   Exit Sub
End If

If Evento = "Novo" Then
    nCodReduz = Val(frmCadMob.txtCodEmpresa.Text)
    Sql = "insert sil (codigo,protocolo,data_emissao,data_validade) values(" & nCodReduz & ",'" & Mask(txtProtocolo.Text) & "','"
    Sql = Sql & Format(mskDataEmissao.Text, "mm/dd/yyyy") & "','" & Format(mskDataValidade.Text, "mm/dd/yyyy") & "')"
Else
    nSid = Val(lblSid.Caption)
    Sql = "update sil set protocolo='" & Mask(txtProtocolo.Text) & "',data_emissao='" & Format(mskDataEmissao.Text, "mm/dd/yyyy") & "',"
    Sql = Sql & "data_validade='" & Format(mskDataValidade.Text, "mm/dd/yyyy") & "' where sid=" & nSid
End If
cn.Execute Sql, rdExecDirect
Limpa
CarregaLista

Evento = ""
Eventos "INICIAR"
If lvMain.ListItems.Count > 0 Then
    lvMain.ListItems(1).Selected = True
    Le
End If

End Sub

Private Sub cmdNovo_Click()

Limpa
Eventos "INCLUIR"
Evento = "Novo"

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim itmX As ListItem
nCodReduz = Val(frmCadMob.txtCodEmpresa.Text)
Me.Top = frmCadMob.Top + 3400
Me.Left = frmCadMob.Left + 7000
CarregaLista
If lvMain.ListItems.Count > 0 Then
    lvMain.ListItems(1).Selected = True
    Le
End If
Evento = "Novo"
Eventos "INICIAR"




End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem, sDataEmissao As String, sDataValidade As String, sArea As String
lvMain.ListItems.Clear
Sql = "select * from sil where codigo=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If IsNull(!Data_Emissao) Then
            sDataEmissao = ""
        Else
            If IsDate(!Data_Emissao) Then
                sDataEmissao = Format(!Data_Emissao, "dd/mm/yyyy")
            Else
                sDataEmissao = ""
            End If
        End If
        If IsNull(!Data_Validade) Then
            sDataValidade = ""
        Else
            If IsDate(!Data_Validade) Then
                sDataValidade = Format(!Data_Validade, "dd/mm/yyyy")
            Else
                sDataValidade = ""
            End If
        End If
                        
                        
        Set itmX = lvMain.ListItems.Add(, , SubNull(!Protocolo))
        itmX.SubItems(1) = sDataEmissao
        itmX.SubItems(2) = sDataValidade
        itmX.SubItems(3) = !sID
       .MoveNext
    Loop
   .Close
End With

End Sub


Private Sub LoadLista()
Dim Sql As String, RdoAux As rdoResultset, sDataEmissao As String, sDataValidade As String
lstSil.Clear
Sql = "select * from sil where codigo=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!Data_Emissao) Then
        sDataEmissao = "__/__/____"
    Else
        If IsDate(!Data_Emissao) Then
            sDataEmissao = Format(!Data_Emissao, "dd/mm/yyyy")
        Else
            sDataEmissao = "__/__/____"
        End If
    End If
    If IsNull(!Data_Validade) Then
        sDataValidade = "__/__/____"
    Else
        If IsDate(!Data_Validade) Then
            sDataValidade = Format(!Data_Validade, "dd/mm/yyyy")
        Else
            sDataValidade = "__/__/____"
        End If
    End If
    Do Until .EOF
        lstSil.AddItem !Sil & " (Emissão: " & sDataEmissao & " - Validade: " & sDataValidade & ")"
        lstSil.ItemData(lstSil.NewIndex) = !sID
       .MoveNext
    Loop
   .Close
End With
If lstSil.ListCount > 0 Then lstSil.ListIndex = 0

End Sub

Private Sub Limpa()
txtProtocolo.Text = ""
LimpaMascara mskDataEmissao
LimpaMascara mskDataValidade
End Sub

Private Sub lvMain_Click()
Limpa
Le
End Sub

Private Sub Le()
If lvMain.ListItems.Count = 0 Then Exit Sub
If lvMain.SelectedItem.Index = -1 Then Exit Sub
txtProtocolo.Text = lvMain.SelectedItem.Text
If lvMain.SelectedItem.SubItems(1) <> "" Then
    mskDataEmissao.Text = lvMain.SelectedItem.SubItems(1)
End If
If lvMain.SelectedItem.SubItems(2) <> "" Then
    mskDataValidade.Text = lvMain.SelectedItem.SubItems(2)
End If
lblSid.Caption = lvMain.SelectedItem.SubItems(3)
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
Limpa
Le
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   txtProtocolo.Enabled = False
   txtProtocolo.BackColor = Me.BackColor
   mskDataEmissao.Enabled = False
   mskDataEmissao.BackColor = Me.BackColor
   mskDataValidade.Enabled = False
   mskDataValidade.BackColor = Me.BackColor
   lvMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   txtProtocolo.Enabled = True
   txtProtocolo.BackColor = Branco
   mskDataEmissao.Enabled = True
   mskDataEmissao.BackColor = Branco
   mskDataValidade.Enabled = True
   mskDataValidade.BackColor = Branco
   lvMain.Enabled = True
End If

If NomeDeLogin <> "RITA" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "SCHWARTZ" Then
    cmdNovo.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
End If

End Sub

