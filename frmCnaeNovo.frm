VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCnaeNovo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista CNAE"
   ClientHeight    =   7065
   ClientLeft      =   10020
   ClientTop       =   2895
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9120
   Begin Tributacao.jcFrames jcFrames 
      Height          =   2535
      Left            =   90
      Top             =   4455
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   4471
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Inclusão/Alteração de CNAE"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   660
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1980
         Width           =   1605
      End
      Begin VB.ComboBox cmbCriterio 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1575
         Width           =   2775
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   1125
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   6180
      End
      Begin esMaskEdit.esMaskedEdit mskCnae 
         Height          =   330
         Left            =   1125
         TabIndex        =   5
         Top             =   270
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         MouseIcon       =   "frmCnaeNovo.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         MaxLength       =   9
         Mask            =   "9999-9/99"
         SelText         =   ""
         Text            =   "____-_/__"
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdGravar 
         Height          =   315
         Left            =   6210
         TabIndex        =   7
         ToolTipText     =   "Gravar os Dados"
         Top             =   225
         Width           =   1065
         _ExtentX        =   1879
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
         MICON           =   "frmCnaeNovo.frx":001C
         PICN            =   "frmCnaeNovo.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid grdTmp 
         Height          =   1125
         Left            =   3465
         TabIndex        =   10
         Top             =   1305
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   1984
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   ">Seq   |>Item   |<Descrição                 |>Valor          "
      End
      Begin prjChameleon.chameleonButton cmdAdd 
         Height          =   315
         Left            =   2925
         TabIndex        =   11
         ToolTipText     =   "Adicionar critério"
         Top             =   1575
         Width           =   495
         _ExtentX        =   873
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnaeNovo.frx":03DD
         PICN            =   "frmCnaeNovo.frx":03F9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDel 
         Height          =   315
         Left            =   2925
         TabIndex        =   12
         ToolTipText     =   "Remover critério"
         Top             =   1935
         Width           =   495
         _ExtentX        =   873
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnaeNovo.frx":0553
         PICN            =   "frmCnaeNovo.frx":056F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor..:"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione os critérios:"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   1305
         Width           =   1920
      End
      Begin VB.Label Label 
         Caption         =   "Descrição..:"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label 
         Caption         =   "CNAE....:"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   360
         Width           =   915
      End
   End
   Begin prjChameleon.chameleonButton cmdSelect 
      Height          =   345
      Left            =   7695
      TabIndex        =   2
      ToolTipText     =   "Pesquisar"
      Top             =   6615
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "S&elecionar"
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
      MICON           =   "frmCnaeNovo.frx":06C9
      PICN            =   "frmCnaeNovo.frx":06E5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPesq 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8970
   End
   Begin MSComctlLib.ListView lvCnae 
      Height          =   3855
      Left            =   45
      TabIndex        =   1
      Top             =   495
      Width           =   9030
      _ExtentX        =   15928
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
         Text            =   "CNAE"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   13759
      EndProperty
   End
End
Attribute VB_Name = "frmCnaeNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NomeForm As String

Public Property Let sForm(sNomeForm As String)
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, sClasse As String, sCnae As String

NomeForm = sNomeForm
If NomeForm = "Menu" Then cmdSelect.Enabled = False
    
End Property

Private Sub cmbCriterio_Click()
If cmbCriterio.ListIndex = -1 Then Exit Sub

Sql = "SELECT VALOR FROM CNAECRITERIODESC WHERE CRITERIO=" & cmbCriterio.ItemData(cmbCriterio.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    txtValor.Text = FormatNumber(!Valor, 4)
   .Close
End With

If Val(txtValor.Text) = 0 Then
    txtValor.Locked = False
    txtValor.BackColor = Branco
Else
    txtValor.Locked = True
    txtValor.BackColor = Kde
End If

End Sub

Private Sub cmdAdd_Click()
Dim x As Integer, sCnae As String

If cmbCriterio.ListIndex = -1 Then
    MsgBox "Selecione o critério.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtValor.Text) = 0 Then
    MsgBox "Digite o valor.", vbCritical, "Atenção"
    Exit Sub
End If

With grdTmp
    For x = 1 To .Rows - 1
        If Val(.TextMatrix(x, 1)) = cmbCriterio.ItemData(cmbCriterio.ListIndex) Then
            MsgBox "Item já incluido na lista.", vbCritical, "Atenção"
            Exit Sub
        End If
    Next
    
    
    sCnae = RetornaNumero(mskCnae.Text)
   .AddItem .Rows & Chr(9) & cmbCriterio.ItemData(cmbCriterio.ListIndex) & Chr(9) & cmbCriterio.Text & Chr(9) & FormatNumber(txtValor.Text, 2)
'    Sql = "DELETE FROM CNAECRITERIO WHERE cnae='" & sCnae & "'"
'    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM cnae_criterio WHERE cnae='" & sCnae & "'"
    cn.Execute Sql, rdExecDirect
    
    
    With grdTmp
        For x = 1 To grdTmp.Rows - 1
'            Sql = "INSERT CNAECRITERIO(CNAE,SEQ,CRITERIO,VALOR) VALUES('" & sCnae & "'," & .TextMatrix(x, 0) & ","
'            Sql = Sql & .TextMatrix(x, 1) & "," & Virg2Ponto(.TextMatrix(x, 3)) & ")"
            Sql = "INSERT cnae_criterio(CNAE,CRITERIO) VALUES('" & sCnae & "'," & .TextMatrix(x, 1) & ")"
            cn.Execute Sql, rdExecDirect
        Next
    End With
End With

End Sub

Private Sub cmdDel_Click()
Dim sCnae As String

With grdTmp
    If .Rows = 1 Then Exit Sub
    If .Rows = 2 Then
       grdTmp.Rows = 1
    Else
        .RemoveItem (.Row)
    End If
End With
 
 sCnae = RetornaNumero(mskCnae.Text)
' Sql = "DELETE FROM CNAECRITERIO WHERE cnae='" & sCnae & "'"
' cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM cnae_criterio_valor WHERE cnae='" & sCnae & "'"
    cn.Execute Sql, rdExecDirect
 
 With grdTmp
     For x = 1 To grdTmp.Rows - 1
'         Sql = "INSERT CNAECRITERIO(CNAE,SEQ,CRITERIO,VALOR) VALUES('" & sCnae & "'," & .TextMatrix(x, 0) & ","
'         Sql = Sql & .TextMatrix(x, 1) & "," & Virg2Ponto(.TextMatrix(x, 3)) & ")"
          Sql = "INSERT cnae_criterio_valor(CNAE,CRITERIO,VALOR) VALUES('" & sCnae & "'," & .TextMatrix(x, 1) & "," & Virg2Ponto(.TextMatrix(x, 3)) & ")"
          cn.Execute Sql, rdExecDirect
          cn.Execute Sql, rdExecDirect
     Next
 End With

End Sub

Private Sub cmdGravar_Click()
Dim bAchou As Boolean, x As Integer, sCnae As String, Sql As String


If NomeDeLogin <> "RITA" And NOMEDLEOGIN <> "LEANDRO" And NomeDeLogin <> "PAULA" And NomeDeLogin <> "RODRIGOC" And NomeDeLogin <> "LUIZH" And NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Permissão negada", vbCritical, "Atenção"
    Exit Sub
End If


sCnae = RetornaNumero(mskCnae.Text)
If Len(sCnae) < 7 Then
    MsgBox "CNAE deve ter 7 dígitos.", vbCritical, "Erro"
    Exit Sub
End If

bAchou = False
For x = 1 To lvCnae.ListItems.Count
    If lvCnae.ListItems(x).Text = sCnae Then
        bAchou = True
        Exit For
    End If
Next

If Trim(txtDesc.Text) = "" Then
    MsgBox "Digite a descrição do CNAE.", vbCritical, "Atenção"
    Exit Sub
End If

If bAchou Then
    If MsgBox("Deseja ALTERAR a descrição do CNAE existente de " & lvCnae.SelectedItem.SubItems(1) & " para " & txtDesc.Text & "?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Sql = "update cnae set descricao='" & Mask(txtDesc.Text) & "' where cnae='" & sCnae & "'"
        cn.Execute Sql, rdExecDirect
        CarregaLista
    End If
Else
    If MsgBox("Deseja incluir este NOVO CNAE?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Sql = "insert cnae(cnae,descricao) values('" & sCnae & "','" & Mask(txtDesc.Text) & "')"
        cn.Execute Sql, rdExecDirect
        CarregaLista
    End If
End If


End Sub

Private Sub cmdSelect_Click()
Dim bAchou As Boolean, x As Integer, sCnae As String, sDesc As String

If lvCnae.ListItems.Count = 0 Then
   MsgBox "Nenhum registro selecionado.", vbInformation, "Atenção"
End If

sCnae = Format(lvCnae.SelectedItem.Text, "0000-0/00")
sDesc = lvCnae.SelectedItem.SubItems(1)

If NomeForm = "frmCadMob" Then
    frmCadMob.mskCnae.Text = sCnae
    Unload Me
ElseIf NomeForm = "frmCadMob1" Then
    If sCnae = frmCadMob.mskCnae.Text Then
        MsgBox "CNAE ja cadastrada.", vbCritical, "Atenção"
        Exit Sub
    End If
    bAchou = False
    For x = 0 To frmCadMob.cmbCnae.ListCount - 1
        If frmCadMob.cmbCnae.Text = sCnae Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        frmCadMob.cmbCnae.AddItem sCnae
        frmCadMob.cmbCnae.ListIndex = frmCadMob.cmbCnae.ListCount - 1
        Unload Me
    Else
        MsgBox "CNAE ja cadastrada.", vbCritical, "Atenção"
    End If
ElseIf NomeForm = "frmCadMob2" Then
    frmCadMob.cmbCnae.RemoveItem (frmCadMob.cmbCnae.ListIndex)
    frmCadMob.cmbCnae.AddItem sCnae
    frmCadMob.cmbCnae.ListIndex = frmCadMob.cmbCnae.ListCount - 1
    Unload Me
End If

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaLista
If NomeDeLogin <> "RITA" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "SCHWARTZ" Then
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
Exit Sub
End If


End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset
Dim itmX As ListItem, z As Long
z = SendMessage(lvCnae.HWND, LVM_DELETEALLITEMS, 0, 0)

Ocupado
Sql = "select * from cnae where 1=1 "
If Trim(txtPesq.Text) <> "" Then
    If IsNumeric(txtPesq.Text) Then
        Sql = Sql & " and cnae like '" & Mask(txtPesq.Text) & "%'"
    Else
        Sql = Sql & " and descricao like '%" & Mask(txtPesq.Text) & "%'"
    End If
End If
Sql = Sql & " order by cnae"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvCnae.ListItems.Add(, , !Cnae)
        itmX.SubItems(1) = !descricao
       .MoveNext
    Loop
   .Close
End With


cmbCriterio.Clear

Sql = "SELECT CRITERIO,DESCRICAO FROM CNAECRITERIODESC ORDER BY DESCRICAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbCriterio.AddItem !descricao
        cmbCriterio.ItemData(cmbCriterio.NewIndex) = !criterio
       .MoveNext
    Loop
   .Close
End With



If lvCnae.ListItems.Count > 0 Then
'    lvCnae.ListItems(1).Selected = True
'    itmX = lvCnae.SelectedItem
    lvCnae_ItemClick lvCnae.SelectedItem
End If

Liberado


End Sub

Private Sub lvCnae_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvCnae.ListItems.Count = 0 Then Exit Sub
mskCnae.Text = Format(lvCnae.SelectedItem.Text, "0000-0/00")
txtDesc.Text = lvCnae.SelectedItem.SubItems(1)

End Sub

Private Sub mskCnae_Change()
If Len(mskCnae.ClipText) = 7 Then
    Le
Else
    grdTmp.Rows = 1
End If
End Sub

Private Sub mskCnae_GotFocus()
mskCnae.SelStart = 0
mskCnae.SelLength = Len(mskCnae.Text)
mskCnae.SetFocus
End Sub

Private Sub txtPesq_Change()
CarregaLista
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub Le()
Dim x As Integer, sCnae
sCnae = RetornaNumero(mskCnae.Text)

grdTmp.Rows = 1
'Sql = "SELECT cnaecriterio.cnae, cnaecriterio.seq,cnaecriterio.criterio,cnaecriteriodesc.descricao , cnaecriterio.valor FROM cnaecriterio INNER JOIN cnaecriteriodesc ON cnaecriterio.criterio = cnaecriteriodesc.criterio "
'Sql = Sql & "WHERE CNAE='" & sCnae & "'"
Sql = "SELECT cnae_criterio.cnae, cnae_criterio.criterio, cnaecriteriodesc.descricao, cnaecriteriodesc.valor "
Sql = Sql & "FROM cnae_criterio INNER JOIN cnaecriteriodesc ON cnae_criterio.criterio = cnaecriteriodesc.criterio "
Sql = Sql & "WHERE cnae_criterio.cnae = '" & sCnae & "'"


Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        grdTmp.AddItem 0 & Chr(9) & RdoAux!criterio & Chr(9) & RdoAux!descricao & Chr(9) & FormatNumber(RdoAux!Valor, 4)
       .MoveNext
    Loop
   .Close
End With

End Sub

