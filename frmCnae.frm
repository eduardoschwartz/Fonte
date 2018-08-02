VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCnae 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de CNAE Fiscal 2.0"
   ClientHeight    =   4500
   ClientLeft      =   4095
   ClientTop       =   3150
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   9165
   Begin MSFlexGridLib.MSFlexGrid grdTmp 
      Height          =   1395
      Left            =   5220
      TabIndex        =   21
      Top             =   2910
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   2461
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   ">Seq   |>Item   |<Descrição                 |>Valor          "
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   690
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3600
      Width           =   1605
   End
   Begin VB.ComboBox cmbCriterio 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3150
      Width           =   4305
   End
   Begin VB.ComboBox cmbSubClasse 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1650
      Width           =   8055
   End
   Begin VB.ComboBox cmbClasse 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1260
      Width           =   8055
   End
   Begin VB.ComboBox cmbGrupo 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   870
      Width           =   8055
   End
   Begin VB.ComboBox cmbDivisao 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   8055
   End
   Begin VB.ComboBox cmbSecao 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   8055
   End
   Begin prjChameleon.chameleonButton cmdSelect 
      Height          =   345
      Left            =   5250
      TabIndex        =   6
      ToolTipText     =   "Pesquisar"
      Top             =   2130
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
      MICON           =   "frmCnae.frx":0000
      PICN            =   "frmCnae.frx":001C
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
      Height          =   345
      Left            =   7830
      TabIndex        =   8
      ToolTipText     =   "Sair da Tela"
      Top             =   2130
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
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
      MICON           =   "frmCnae.frx":0176
      PICN            =   "frmCnae.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskCnae 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2100
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      MouseIcon       =   "frmCnae.frx":0200
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
   Begin prjChameleon.chameleonButton cmdCriterio 
      Default         =   -1  'True
      Height          =   345
      Left            =   6540
      TabIndex        =   7
      ToolTipText     =   "Pesquisar"
      Top             =   2130
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "C&ritérios"
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
      MICON           =   "frmCnae.frx":021C
      PICN            =   "frmCnae.frx":0238
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAdd 
      Height          =   315
      Left            =   4500
      TabIndex        =   11
      ToolTipText     =   "Adicionar critério"
      Top             =   3150
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
      MICON           =   "frmCnae.frx":02B4
      PICN            =   "frmCnae.frx":02D0
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
      Left            =   4500
      TabIndex        =   12
      ToolTipText     =   "Remover critério"
      Top             =   3540
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
      MICON           =   "frmCnae.frx":042A
      PICN            =   "frmCnae.frx":0446
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
      Left            =   120
      TabIndex        =   20
      Top             =   3660
      Width           =   675
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione os critérios:"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   2910
      Width           =   4035
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   90
      X2              =   9090
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Código CNAE..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   18
      Top             =   2130
      Width           =   1905
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SubClasse..:"
      Height          =   225
      Left            =   60
      TabIndex        =   17
      Top             =   1710
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe........:"
      Height          =   225
      Left            =   60
      TabIndex        =   16
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo........:"
      Height          =   225
      Left            =   60
      TabIndex        =   15
      Top             =   930
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Divisão......:"
      Height          =   225
      Left            =   60
      TabIndex        =   14
      Top             =   540
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seção.......:"
      Height          =   225
      Left            =   60
      TabIndex        =   13
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "frmCnae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, bExec As Boolean, NomeForm As String
Dim sSecao As String, nDivisao As Integer, nGrupo As Integer, nClasse As Integer, nSubClasse As Integer

Public Property Let sForm(sNomeForm As String)
Dim a As Integer, B As Integer, c As Integer, d As Integer, e As Integer, sClasse As String, sCnae As String

NomeForm = sNomeForm
If NomeForm = "Menu" Then cmdSelect.Enabled = False

If NomeForm <> "frmCadMob" And NomeForm <> "frmCadMob2" Then Exit Property

If NomeForm = "frmCadMob" And frmCadMob.cmdGravar.Visible = False Then cmdSelect.Enabled = False
If NomeForm = "frmCadMob" Then
    If frmCadMob.mskCnae.ClipText <> "" Then
        sCnae = frmCadMob.mskCnae.Text
    End If
ElseIf NomeForm = "frmCadMob2" Then
    sCnae = frmCadMob.cmbCnae.Text
End If

sSecao = Left(sCnae, 1)
For a = 0 To cmbSecao.ListCount - 1
    If Left(cmbSecao.List(a), 1) = sSecao Then
        cmbSecao.ListIndex = a
        nDivisao = Val(Mid(sCnae, 2, 2))
        For B = 0 To cmbDivisao.ListCount - 1
            If Val(Left(cmbDivisao.List(B), 2)) = nDivisao Then
                cmbDivisao.ListIndex = B
                nGrupo = Val(Mid(sCnae, 4, 1))
                For c = 0 To cmbGrupo.ListCount - 1
                    If Val(Left(cmbGrupo.List(c), 1)) = nGrupo Then
                        cmbGrupo.ListIndex = c
                        sClasse = Mid(sCnae, 5, 3)
                        sClasse = Left(sClasse, 1) & Right(sClasse, 1)
                        nClasse = Val(sClasse)
                        For d = 0 To cmbClasse.ListCount - 1
                            If Val(Left(cmbClasse.List(d), 2)) = nClasse Then
                                cmbClasse.ListIndex = d
                                nSubClasse = Val(Right(sCnae, 1))
                                For e = 0 To cmbSubClasse.ListCount - 1
                                    If Val(Left(cmbSubClasse.List(e), 2)) = nSubClasse Then
                                        cmbSubClasse.ListIndex = e
                                        Exit For
                                    End If
                                Next e
                                Exit For
                            End If
                        Next d
                        Exit For
                    End If
                Next c
                Exit For
            End If
        Next B
        Exit For
    End If
Next a
    
End Property


Private Sub cmbCriterio_Click()
If cmbCriterio.ListIndex = -1 Then Exit Sub

Sql = "SELECT VALOR FROM CNAECRITERIODESC WHERE CRITERIO=" & cmbCriterio.ItemData(cmbCriterio.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    txtValor.Text = FormatNumber(!Valor, 2)
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

Private Sub cmbDivisao_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbSecao_Click()
If Not bExec Then Exit Sub
If cmbSecao.ListIndex = -1 Then Exit Sub
bExec = False
'LimpaMascara mskCnae
CarregaDivisao
bExec = True
End Sub

Private Sub cmbDivisao_Click()
If Not bExec Then Exit Sub
If cmbDivisao.ListIndex = -1 Then Exit Sub
bExec = False
CarregaGrupo
bExec = True
End Sub

Private Sub cmbGRUPO_Click()
If Not bExec Then Exit Sub
If cmbGrupo.ListIndex = -1 Then Exit Sub
bExec = False
CarregaClasse
bExec = True
End Sub

Private Sub cmbClasse_Click()
If Not bExec Then Exit Sub
If cmbClasse.ListIndex = -1 Then Exit Sub
bExec = False
CarregaSubClasse
bExec = True
End Sub

Private Sub cmdAdd_Click()
Dim x As Integer

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
   .AddItem .Rows & Chr(9) & cmbCriterio.ItemData(cmbCriterio.ListIndex) & Chr(9) & cmbCriterio.Text & Chr(9) & FormatNumber(txtValor.Text, 2)
End With

End Sub

Private Sub cmdCriterio_Click()
Dim Ct As Control

If cmbSubClasse.ListIndex = -1 Then
    MsgBox "Selecione um Código CNAE válido.", vbCritical, "Atenção"
    cmdCriterio.Value = False
    Exit Sub
End If
If cmdCriterio.Value = False Then
    Me.Height = 3120
Else
    Me.Height = 4965
End If
Centraliza Me
If cmdCriterio.Value = True Then
    For Each Ct In frmCnae
        If TypeOf Ct Is esMaskedEdit Or TypeOf Ct Is ComboBox Then
           Ct.BackColor = Kde
           Ct.Enabled = False
        End If
    Next
    cmdSelect.Enabled = False
    cmdSair.Enabled = False
    CarregaCriterio
    Le
Else
    For Each Ct In frmCnae
        If TypeOf Ct Is esMaskedEdit Or TypeOf Ct Is ComboBox Then
           Ct.BackColor = Branco
           Ct.Enabled = True
        End If
    Next
    cmdSelect.Enabled = True
    cmdSair.Enabled = True
    Grava
End If
cmbCriterio.BackColor = Branco
cmbCriterio.Enabled = True

End Sub

Private Sub cmdDel_Click()
With grdTmp
    If .Rows = 1 Then Exit Sub
    If .Rows = 2 Then
       grdTmp.Rows = 1
    Else
        .RemoveItem (.Row)
    End If
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSelect_Click()
Dim bAchou As Boolean, x As Integer

If cmbSubClasse.ListIndex = -1 Then
    MsgBox "Selecione o código completo.", vbExclamation, "Atenção"
    Exit Sub
End If
CodEmpresa = 0
If NomeForm = "frmCadMob" Then
    frmCadMob.mskCnae.Text = mskCnae.Text
    Unload Me
ElseIf NomeForm = "frmCadMob1" Then
    If mskCnae.Text = frmCadMob.mskCnae.Text Then
        MsgBox "CNAE ja cadastrada.", vbCritical, "Atenção"
        Exit Sub
    End If
    bAchou = False
    For x = 0 To frmCadMob.cmbCnae.ListCount - 1
        If frmCadMob.cmbCnae.Text = mskCnae.Text Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        frmCadMob.cmbCnae.AddItem mskCnae.Text
        frmCadMob.cmbCnae.ListIndex = frmCadMob.cmbCnae.ListCount - 1
        Unload Me
    Else
        MsgBox "CNAE ja cadastrada.", vbCritical, "Atenção"
    End If
ElseIf NomeForm = "frmCadMob2" Then
    frmCadMob.cmbCnae.RemoveItem (frmCadMob.cmbCnae.ListIndex)
    frmCadMob.cmbCnae.AddItem mskCnae.Text
    frmCadMob.cmbCnae.ListIndex = frmCadMob.cmbCnae.ListCount - 1
    Unload Me
End If


End Sub

Private Sub Form_Load()
Centraliza Me
bExec = False
CarregaSecao
bExec = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Grava
End Sub

Private Sub CarregaSecao()
cmbSecao.Clear: cmbDivisao.Clear: cmbGrupo.Clear: cmbClasse.Clear: cmbSubClasse.Clear
Sql = "SELECT SECAO,DESCRICAO FROM CNAESECAO ORDER BY SECAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbSecao.AddItem !secao & " - " & !Descricao
       .MoveNext
    Loop
   .Close
End With
End Sub

Private Sub CarregaDivisao()
cmbDivisao.Clear: cmbGrupo.Clear: cmbClasse.Clear: cmbSubClasse.Clear
Sql = "SELECT SECAO,DIVISAO,DESCRICAO FROM CNAEDIVISAO WHERE SECAO='" & Left$(cmbSecao.Text, 1) & "' ORDER BY DIVISAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbDivisao.AddItem Format(!divisao, "00") & " - " & !Descricao
       .MoveNext
    Loop
   .Close
End With
End Sub

Private Sub CarregaGrupo()
cmbGrupo.Clear: cmbClasse.Clear: cmbSubClasse.Clear
Sql = "SELECT SECAO,DIVISAO,GRUPO,DESCRICAO FROM CNAEGRUPO WHERE SECAO='" & Left$(cmbSecao.Text, 1) & "' AND DIVISAO=" & Val(Left$(cmbDivisao.Text, 2)) & " ORDER BY GRUPO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbGrupo.AddItem !grupo & " - " & !Descricao
       .MoveNext
    Loop
   .Close
End With
End Sub

Private Sub CarregaClasse()
cmbClasse.Clear: cmbSubClasse.Clear
Sql = "SELECT SECAO,DIVISAO,GRUPO,CLASSE,DESCRICAO FROM CNAECLASSE WHERE SECAO='" & Left$(cmbSecao.Text, 1) & "' AND DIVISAO=" & Val(Left$(cmbDivisao.Text, 2)) & "AND GRUPO=" & Val(Left$(cmbGrupo.Text, 1)) & " ORDER BY CLASSE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbClasse.AddItem Format(!classe, "00") & " - " & !Descricao
       .MoveNext
    Loop
   .Close
End With
End Sub

Private Sub CarregaSubClasse()
cmbSubClasse.Clear
Sql = "SELECT SECAO,DIVISAO,GRUPO,CLASSE,SUBCLASSE,DESCRICAO FROM CNAESUBCLASSE WHERE DIVISAO=" & Val(Left$(cmbDivisao.Text, 2)) & "AND GRUPO=" & Val(Left$(cmbGrupo.Text, 1)) & "AND CLASSE=" & Val(Left$(cmbClasse.Text, 2)) & " ORDER BY SUBCLASSE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbSubClasse.AddItem Format(!subclasse, "00") & " - " & !Descricao
       .MoveNext
    Loop
   .Close
End With
End Sub

Private Sub MontaCodigo()
'Dim sClasse As String, sCnae As String
'If cmbDivisao.ListIndex > -1 Then
'    sCnae = Left$(cmbDivisao.text, 2)
'    If cmbGrupo.ListIndex > -1 Then
'        sCnae = sCnae & Left$(cmbGrupo.text, 1)
'        If cmbClasse.ListIndex > -1 Then
'            sClasse = Left$(cmbClasse.text, 2)
'            sCnae = sCnae & Left$(sClasse, 1) & "-" & Right$(sClasse, 1)
'            If cmbSubClasse.ListIndex > -1 Then
'            sCnae = sCnae & "/" & Left$(cmbSubClasse.text, 2)
'            End If
'        End If
'    End If
'End If
'bExec = False
'mskCnae.text = sCnae
'bExec = True
End Sub

Private Sub mskCnae_Change()
Dim sTexto As String, sDivisao As String, x As Integer, sGrupo As String, sClasse As String, sSubClasse As String
If Not bExec Then Exit Sub
sTexto = mskCnae.ClipText
'If sTexto = "" Then Exit Sub
If Len(sTexto) < 2 Then
    cmbSecao.ListIndex = -1
    cmbDivisao.Clear: cmbGrupo.Clear: cmbClasse.Clear: cmbSubClasse.Clear
    Exit Sub
End If

sDivisao = Left(sTexto, 2)
Select Case sDivisao
    Case "01", "02", "03"
        cmbSecao.ListIndex = 0
    Case "05", "06", "07", "08", "09"
        cmbSecao.ListIndex = 1
    Case "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33"
        cmbSecao.ListIndex = 2
    Case "35"
        cmbSecao.ListIndex = 3
    Case "36", "37", "38", "39"
        cmbSecao.ListIndex = 4
    Case "41", "42", "43"
        cmbSecao.ListIndex = 5
    Case "45", "46", "47", "50"
        cmbSecao.ListIndex = 6
    Case "49", "51", "52", "53"
        cmbSecao.ListIndex = 7
    Case "55", "56"
        cmbSecao.ListIndex = 8
    Case "58", "59", "60", "61", "62", "63"
        cmbSecao.ListIndex = 9
    Case "64", "65", "66"
        cmbSecao.ListIndex = 10
    Case "68"
        cmbSecao.ListIndex = 11
    Case "69", "70", "71", "72", "73", "74", "75"
        cmbSecao.ListIndex = 12
    Case "77", "78", "79", "80", "81", "82"
        cmbSecao.ListIndex = 13
    Case "84"
        cmbSecao.ListIndex = 14
    Case "85"
        cmbSecao.ListIndex = 15
    Case "86", "87", "88"
        cmbSecao.ListIndex = 16
    Case "90", "91", "92", "93"
        cmbSecao.ListIndex = 17
    Case "94", "95", "96"
        cmbSecao.ListIndex = 18
    Case "97"
        cmbSecao.ListIndex = 19
    Case "99"
        cmbSecao.ListIndex = 20
    Case Else
        cmbSecao.ListIndex = -1
        Exit Sub
End Select

For x = 0 To cmbDivisao.ListCount - 1
    If Left(cmbDivisao.List(x), 2) = sDivisao Then
        cmbDivisao.ListIndex = x
        Exit For
    End If
Next

If Len(sTexto) < 3 Then Exit Sub
sGrupo = Mid(sTexto, 3, 1)
cmbGrupo.ListIndex = -1
For x = 0 To cmbGrupo.ListCount - 1
    If Left(cmbGrupo.List(x), 1) = sGrupo Then
        cmbGrupo.ListIndex = x
        Exit For
    End If
Next

If Len(sTexto) < 5 Then Exit Sub
sClasse = Mid(sTexto, 4, 2)
cmbClasse.ListIndex = -1
For x = 0 To cmbClasse.ListCount - 1
    If Left(cmbClasse.List(x), 2) = sClasse Then
        cmbClasse.ListIndex = x
        Exit For
    End If
Next

If Len(sTexto) < 7 Then Exit Sub
sSubClasse = Mid(sTexto, 6, 2)
cmbSubClasse.ListIndex = -1
For x = 0 To cmbSubClasse.ListCount - 1
    If Left(cmbSubClasse.List(x), 2) = sSubClasse Then
        cmbSubClasse.ListIndex = x
        Exit For
    End If
Next

End Sub

Private Sub mskCnae_GotFocus()
mskCnae.SetFocus
End Sub

Private Sub CarregaCriterio()
cmbCriterio.Clear

Sql = "SELECT CRITERIO,DESCRICAO FROM CNAECRITERIODESC ORDER BY DESCRICAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbCriterio.AddItem !Descricao
        cmbCriterio.ItemData(cmbCriterio.NewIndex) = !CRITERIO
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub Grava()
Dim x As Integer, sSecao As String, nDivisao As Integer, nGrupo As Integer, nClasse As Integer, nSubClasse As Integer
sSecao = Left$(cmbSecao.Text, 1)
nDivisao = Val(Left$(cmbDivisao.Text, 2))
nGrupo = Val(Left$(cmbGrupo.Text, 1))
nClasse = Val(Left$(cmbClasse.Text, 2))
nSubClasse = Val(Left$(cmbSubClasse.Text, 2))

Sql = "DELETE FROM CNAECRITERIO WHERE SECAO='" & sSecao & "' AND DIVISAO=" & nDivisao & " AND "
Sql = Sql & "GRUPO=" & nGrupo & " AND CLASSE=" & nClasse & " AND SUBCLASSE=" & nSubClasse
cn.Execute Sql, rdExecDirect

With grdTmp
    For x = 1 To grdTmp.Rows - 1
        Sql = "INSERT CNAECRITERIO(SECAO,DIVISAO,GRUPO,CLASSE,SUBCLASSE,SEQ,CRITERIO,VALOR) VALUES('" & sSecao & "',"
        Sql = Sql & nDivisao & "," & nGrupo & "," & nClasse & "," & nSubClasse & "," & .TextMatrix(x, 0) & ","
        Sql = Sql & .TextMatrix(x, 1) & "," & Virg2Ponto(.TextMatrix(x, 3)) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

End Sub

Private Sub Le()
Dim x As Integer, sSecao As String, nDivisao As Integer, nGrupo As Integer, nClasse As Integer, nSubClasse As Integer
sSecao = Left$(cmbSecao.Text, 1)
nDivisao = Val(Left$(cmbDivisao.Text, 2))
nGrupo = Val(Left$(cmbGrupo.Text, 1))
nClasse = Val(Left$(cmbClasse.Text, 2))
nSubClasse = Val(Left$(cmbSubClasse.Text, 2))

grdTmp.Rows = 1
Sql = "SELECT cnaecriterio.secao, cnaecriterio.divisao, cnaecriterio.grupo, cnaecriterio.classe, cnaecriterio.subclasse, cnaecriterio.seq, cnaecriterio.criterio, "
Sql = Sql & "cnaecriteriodesc.descricao , cnaecriterio.valor FROM cnaecriterio INNER JOIN cnaecriteriodesc ON cnaecriterio.criterio = cnaecriteriodesc.criterio "
Sql = Sql & "WHERE SECAO='" & sSecao & "' AND DIVISAO=" & nDivisao & " AND GRUPO=" & nGrupo & " AND CLASSE=" & nClasse & " AND SUBCLASSE=" & nSubClasse
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        grdTmp.AddItem !Seq & Chr(9) & !CRITERIO & Chr(9) & !Descricao & Chr(9) & FormatNumber(!Valor, 2)
       .MoveNext
    Loop
   .Close
End With

End Sub
