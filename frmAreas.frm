VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmAreas 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Áreas"
   ClientHeight    =   2430
   ClientLeft      =   5805
   ClientTop       =   4185
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4230
      MaxLength       =   10
      TabIndex        =   1
      Top             =   135
      Width           =   1410
   End
   Begin VB.TextBox txtHist 
      Appearance      =   0  'Flat
      Height          =   1905
      Left            =   30
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3000
      Width           =   5685
   End
   Begin VB.TextBox txtQtdePav 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   255
      Left            =   5070
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   555
   End
   Begin VB.TextBox txtSeq 
      Height          =   315
      Left            =   90
      TabIndex        =   12
      Top             =   4230
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.ComboBox cmbCategConstr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   3435
   End
   Begin VB.ComboBox cmbTipoConstr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   1845
   End
   Begin VB.ComboBox cmbUsoConstr 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtAreaConstr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1950
      TabIndex        =   2
      Text            =   "0,00"
      Top             =   450
      Width           =   915
   End
   Begin esMaskEdit.esMaskedEdit mskDataAprova 
      Height          =   285
      Left            =   1950
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmAreas.frx":0000
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
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4545
      TabIndex        =   16
      ToolTipText     =   "Cancelar Edição"
      Top             =   2010
      Width           =   1080
      _ExtentX        =   1905
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
      MICON           =   "frmAreas.frx":001C
      PICN            =   "frmAreas.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRetorna 
      Height          =   315
      Left            =   3285
      TabIndex        =   17
      ToolTipText     =   "Cadastra a Área"
      Top             =   2010
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cadastrar"
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
      MICON           =   "frmAreas.frx":0192
      PICN            =   "frmAreas.frx":01AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Guia de Multa:"
      Height          =   225
      Index           =   6
      Left            =   3120
      TabIndex        =   18
      Top             =   180
      Width           =   1050
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Histórico:"
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2730
      Width           =   765
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Pavimentos..:"
      Height          =   225
      Index           =   7
      Left            =   3990
      TabIndex        =   13
      Top             =   1230
      Width           =   1035
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Data 1ª Aprovação.......:"
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   195
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Categ. da Construção....:"
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo da Construção.......:"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Uso da Construção........:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Área da Construção.......:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   540
      Width           =   1815
   End
End
Attribute VB_Name = "frmAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String, sTipoEvento As String
Dim NodX As Object, nCont As Integer, sArea As String, nSeqArea As Integer
Dim sTipoUso As String, sTipoTipo As String, sTipoCat As String, nQtdePav As Integer
Dim dDataConst As Date, nAreaConst As Double, sNumProc As String, dDataProc As Date
Dim NomeForm As String, sHistorico As String

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property

Public Property Let sTipoarea(sValue As String)
    sArea = sValue
End Property

Public Property Let nSequenciaArea(nValue As Integer)
    nSeqArea = nValue
End Property

Public Property Let sEvento(sValue As String)
    sTipoEvento = sValue
End Property

Public Property Let sUso(sValue As String)
    sTipoUso = sValue
End Property

Public Property Let sTipo(sValue As String)
    sTipoTipo = sValue
End Property

Public Property Let sCat(sValue As String)
    sTipoCat = sValue
End Property

Public Property Let sHist(sValue As String)
    sHistorico = sValue
End Property

Public Property Let dDataConstrucao(dValue As Date)
    dDataConst = dValue
End Property

Public Property Let dDataProcesso(dValue As Date)
    dDataProc = dValue
End Property

Public Property Let sNumProcesso(dValue As String)
    sNumProc = dValue
End Property

Public Property Let nAreaConstrucao(nValue As Double)
    nAreaConst = nValue
End Property

Public Property Let nQtdePavimento(nValue As Integer)
    nQtdePav = nValue
End Property

Private Sub cmbTipoConstr_Click()
    
If cmbTipoConstr.ListIndex = 0 Then
    txtQtdePav.Locked = True
    txtQtdePav.BackColor = Kde
    txtQtdePav.Text = 1
Else
    txtQtdePav.Locked = False
    txtQtdePav.BackColor = Branco
End If
    
cmbCategConstr.Clear
Sql = "SELECT DISTINCT categconstr.codcategconstr, categconstr.desccategconstr, fatorcateg.coduso, fatorcateg.codtipo, fatorcateg.codcateg, fatorcateg.anocateg "
Sql = Sql & "FROM categconstr INNER JOIN fatorcateg ON categconstr.codcategconstr = fatorcateg.codcateg "
Sql = Sql & "Where FATORCATEG.anocateg = " & Year(Now) & " And FATORCATEG.coduso = " & cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex) & " And "
Sql = Sql & "FATORCATEG.codtipo = " & cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbCategConstr.AddItem !desccategconstr
       cmbCategConstr.ItemData(cmbCategConstr.NewIndex) = !CODCATEGCONSTR
      .MoveNext
    Loop
   .Close
End With
cmbCategConstr.ListIndex = 0
End Sub

Private Sub cmbUsoConstr_Click()

cmbTipoConstr.Clear
cmbTipoConstr.AddItem "TÉRREA"
cmbTipoConstr.ItemData(cmbTipoConstr.NewIndex) = 1
'If cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex) <> 2 Then 'INDUSTRIA SÓ TERREA
    cmbTipoConstr.AddItem "ASSOBRADADO"
    cmbTipoConstr.ItemData(cmbTipoConstr.NewIndex) = 2
'End If
cmbTipoConstr.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRetorna_Click()
Dim RdoAux As rdoResultset, Sql As String, nCodReduz As Long, nSeq As Integer, sObs As String
Dim itmX As ListItem

If Val(txtQtdePav.Text) = 0 Then
   MsgBox "Digite a qtde de pavimentos.", vbExclamation, "Atenção"
   txtQtdePav.SetFocus
   Exit Sub
End If
If Val(txtQtdePav.Text) < 2 And cmbTipoConstr.ListIndex = 1 Then
   MsgBox "Digite a qtde de pavimentos para construção assobradada.", vbExclamation, "Atenção"
   txtQtdePav.SetFocus
   Exit Sub
End If

If Val(txtAreaConstr.Text) = 0 Then
   MsgBox "Digite a Área da Construção.", vbExclamation, "Atenção"
   txtAreaConstr.SetFocus
   Exit Sub
End If

If cmbUsoConstr.ListIndex = -1 Then
   MsgBox "Selecione o Uso da Construção.", vbExclamation, "Atenção"
   cmbUsoConstr.SetFocus
   Exit Sub
End If

If cmbTipoConstr.ListIndex = -1 Then
   MsgBox "Selecione o Tipo da Construção.", vbExclamation, "Atenção"
   cmbTipoConstr.SetFocus
   Exit Sub
End If

If cmbCategConstr.ListIndex = -1 Then
   MsgBox "Selecione uma Categoria de Construção.", vbExclamation, "Atenção"
   cmbCategConstr.SetFocus
   Exit Sub
End If

If sTipoEvento = "Novo" Then
    If Not IsDate(mskDataAprova.Text) Then
        If Val(txtDoc.Text) = 0 Then
            MsgBox "Data da 1ª aprovação ou nº da guia de multa obrigatórios.", vbExclamation, "Atenção"
            Exit Sub
        Else
            Sql = "SELECT parceladocumento.codreduzido FROM parceladocumento INNER JOIN debitotributo ON parceladocumento.codreduzido = debitotributo.codreduzido AND parceladocumento.anoexercicio = debitotributo.anoexercicio AND "
            Sql = Sql & "parceladocumento.codlancamento = debitotributo.codlancamento AND parceladocumento.seqlancamento = debitotributo.seqlancamento AND parceladocumento.NumParcela = debitotributo.NumParcela And "
            Sql = Sql & "parceladocumento.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO Where (parceladocumento.NumDocumento = " & Val(txtDoc.Text) & ") And (debitotributo.CodTributo = 663)"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount = 0 Then
                   .Close
                    MsgBox "Número da guia não cadastrado ou não é uma guia de multa por falta de recolhimento." & vbcrfl & vbCrLf & "Obs: Digite o número da guia sem o dígito verificador.", vbExclamation, "Atenção"
                    Exit Sub
                Else
                    nCodReduz = Val(Left$(frmCadImob.lblCodReduz.Caption, 7))
                    'grava historico
                    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        If IsNull(!maximo) Then
                            nSeq = 1
                        Else
                            nSeq = !maximo + 1
                        End If
                       .Close
                    End With
                    If sArea = "P" Then
                        sObs = "Área principal "
                    Else
                        sObs = "Área complementar "
                    End If
                    sObs = sObs & "seq: " & txtSeq.Text & " cadastrada através de guia nº " & txtDoc.Text & "-" & RetornaDVNumDoc(Val(txtDoc.Text))
                    
'                    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'                    Sql = Sql & nCodReduz & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObs & "','" & NomeDeLogin & "','" & Format(Now, "mm/dd/yyyy") & "')"
                    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
                    Sql = Sql & nCodReduz & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                    cn.Execute Sql, rdExecDirect
                    frmCadImob.grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & nSeq & Chr(9) & sObs & Chr(9) & "GTI" & Chr(9) & Format(Now, "dd/mm/yyyy")
                End If
               .Close
            End With
        End If
    End If
End If

If NomeForm = "frmCadImob" Then
    If sTipoEvento = "Novo" Then
        nSeq = frmCadImob.lvArea.ListItems.Count + 1
        Set itmX = frmCadImob.lvArea.ListItems.Add(, "A" & Format(nSeq, "00"), Format(nSeq, "00"))
        itmX.SubItems(1) = FormatNumber(txtAreaConstr.Text, 2) & " m²"
        itmX.SubItems(2) = mskDataAprova.Text
        itmX.SubItems(3) = cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex)
        itmX.SubItems(4) = cmbUsoConstr.Text
        itmX.SubItems(5) = cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex)
        itmX.SubItems(6) = cmbTipoConstr.Text
        itmX.SubItems(7) = cmbCategConstr.ItemData(cmbCategConstr.ListIndex)
        itmX.SubItems(8) = cmbCategConstr.Text
        itmX.SubItems(9) = txtQtdePav.Text
    Else
        frmCadImob.lvArea.SelectedItem.SubItems(1) = FormatNumber(txtAreaConstr.Text, 2) & " m²"
        frmCadImob.lvArea.SelectedItem.SubItems(2) = mskDataAprova.Text
        frmCadImob.lvArea.SelectedItem.SubItems(3) = cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex)
        frmCadImob.lvArea.SelectedItem.SubItems(4) = cmbUsoConstr.Text
        frmCadImob.lvArea.SelectedItem.SubItems(5) = cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex)
        frmCadImob.lvArea.SelectedItem.SubItems(6) = cmbTipoConstr.Text
        frmCadImob.lvArea.SelectedItem.SubItems(7) = cmbCategConstr.ItemData(cmbCategConstr.ListIndex)
        frmCadImob.lvArea.SelectedItem.SubItems(8) = cmbCategConstr.Text
        frmCadImob.lvArea.SelectedItem.SubItems(9) = txtQtdePav.Text
    End If
    
ElseIf NomeForm = "frmDesmembramento" Then
    If sTipoEvento = "Novo" Then
        nSeq = frmDesmembramento.lvArea.ListItems.Count + 1
        Set itmX = frmDesmembramento.lvArea.ListItems.Add(, "A" & Format(nSeq, "00"), Format(nSeq, "00"))
        itmX.SubItems(1) = FormatNumber(txtAreaConstr.Text, 2) & " m²"
        itmX.SubItems(2) = mskDataAprova.Text
        itmX.SubItems(3) = cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex)
        itmX.SubItems(4) = cmbUsoConstr.Text
        itmX.SubItems(5) = cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex)
        itmX.SubItems(6) = cmbTipoConstr.Text
        itmX.SubItems(7) = cmbCategConstr.ItemData(cmbCategConstr.ListIndex)
        itmX.SubItems(8) = cmbCategConstr.Text
        itmX.SubItems(9) = txtQtdePav.Text
    Else
        frmDesmembramento.lvArea.SelectedItem.SubItems(1) = FormatNumber(txtAreaConstr.Text, 2) & " m²"
        frmDesmembramento.lvArea.SelectedItem.SubItems(2) = mskDataAprova.Text
        frmDesmembramento.lvArea.SelectedItem.SubItems(3) = cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex)
        frmDesmembramento.lvArea.SelectedItem.SubItems(4) = cmbUsoConstr.Text
        frmDesmembramento.lvArea.SelectedItem.SubItems(5) = cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex)
        frmDesmembramento.lvArea.SelectedItem.SubItems(6) = cmbTipoConstr.Text
        frmDesmembramento.lvArea.SelectedItem.SubItems(7) = cmbCategConstr.ItemData(cmbCategConstr.ListIndex)
        frmDesmembramento.lvArea.SelectedItem.SubItems(8) = cmbCategConstr.Text
        frmDesmembramento.lvArea.SelectedItem.SubItems(9) = txtQtdePav.Text
    End If
        
        
Else
    If sTipoEvento = "Novo" Then
        nSeq = frmCadCondominio.lvArea.ListItems.Count + 1
        Set itmX = frmCadCondominio.lvArea.ListItems.Add(, "A" & Format(nSeq, "00"), Format(nSeq, "00"))
        itmX.SubItems(1) = FormatNumber(txtAreaConstr.Text, 2) & " m²"
        itmX.SubItems(2) = mskDataAprova.Text
        itmX.SubItems(3) = cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex)
        itmX.SubItems(4) = cmbUsoConstr.Text
        itmX.SubItems(5) = cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex)
        itmX.SubItems(6) = cmbTipoConstr.Text
        itmX.SubItems(7) = cmbCategConstr.ItemData(cmbCategConstr.ListIndex)
        itmX.SubItems(8) = cmbCategConstr.Text
        itmX.SubItems(9) = txtQtdePav.Text
    Else
        frmCadCondominio.lvArea.SelectedItem.SubItems(1) = FormatNumber(txtAreaConstr.Text, 2) & " m²"
        frmCadCondominio.lvArea.SelectedItem.SubItems(2) = mskDataAprova.Text
        frmCadCondominio.lvArea.SelectedItem.SubItems(3) = cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex)
        frmCadCondominio.lvArea.SelectedItem.SubItems(4) = cmbUsoConstr.Text
        frmCadCondominio.lvArea.SelectedItem.SubItems(5) = cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex)
        frmCadCondominio.lvArea.SelectedItem.SubItems(6) = cmbTipoConstr.Text
        frmCadCondominio.lvArea.SelectedItem.SubItems(7) = cmbCategConstr.ItemData(cmbCategConstr.ListIndex)
        frmCadCondominio.lvArea.SelectedItem.SubItems(8) = cmbCategConstr.Text
        frmCadCondominio.lvArea.SelectedItem.SubItems(9) = txtQtdePav.Text
    End If
        
        
End If

For x = 0 To Forms.Count - 1
    If NomeForm = "frmDesmembramento" Then
          Unload frmAreas
          frmDesmembramento.SetFocus
    ElseIf NomeForm = "frmCadCondominio" Then
          Unload frmAreas
          frmCadCondominio.SetFocus
    End If
Next

Unload Me
End Sub

Private Sub Form_Activate()
If Me.Height < 2800 Then Me.Height = 2850
End Sub

Private Sub Form_Load()
Dim nSeq As Integer, nCont As Integer

Ocupado

Me.Left = frmMdi.ScaleWidth / 2 - Me.Width / 2
Me.Top = frmMdi.ScaleHeight / 2 - Me.Height / 2

If sTipoEvento <> "Alterar" Then
    txtSeq.Text = nSeqArea
End If

If sArea = "P" Then
   Me.Caption = "Cadastro de Área Principal"
Else
   Me.Caption = "Cadastro de Área Complementar"
End If


Sql = "SELECT CODUSOCONSTR,DESCUSOCONSTR FROM USOCONSTR ORDER BY DESCUSOCONSTR"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbUsoConstr.AddItem !descusoconstr
       cmbUsoConstr.ItemData(cmbUsoConstr.NewIndex) = !CODUSOCONSTR
      .MoveNext
    Loop
   .Close
End With
cmbUsoConstr.ListIndex = 0



If sTipoEvento = "Alterar" Then
   If Year(dDataConst) > 1940 Then
      mskDataAprova.Text = Format(dDataConst, "dd/mm/yyyy")
   Else
      LimpaMascara mskDataAprova
   End If
   txtAreaConstr.Text = nAreaConst
   For x = 0 To cmbUsoConstr.ListCount - 1
       cmbUsoConstr.ListIndex = x
       If cmbUsoConstr.ItemData(cmbUsoConstr.ListIndex) = sTipoUso Then
          Exit For
       End If
   Next
   For x = 0 To cmbTipoConstr.ListCount - 1
       cmbTipoConstr.ListIndex = x
       If cmbTipoConstr.ItemData(cmbTipoConstr.ListIndex) = sTipoTipo Then
          Exit For
       End If
   Next
   For x = 0 To cmbCategConstr.ListCount - 1
       cmbCategConstr.ListIndex = x
       If cmbCategConstr.ItemData(cmbCategConstr.ListIndex) = sTipoCat Then
          Exit For
       End If
   Next
   txtQtdePav.Text = nQtdePav
End If

Liberado
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If NomeForm = "frmDesmembramento" Then
    frmDesmembramento.Refresh
ElseIf NomeForm = "frmCadCondominio" Then
    frmCadCondominio.Refresh
Else
    frmCadImob.Refresh
End If
End Sub

Private Sub txtAreaConstr_GotFocus()
txtAreaConstr.SelStart = 0
txtAreaConstr.SelLength = Len(txtAreaConstr.Text)

End Sub

Private Sub txtAreaConstr_KeyPress(KeyAscii As Integer)
Tweak txtAreaConstr, KeyAscii, DecimalPositive
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
Tweak txtDoc, KeyAscii, IntegerPositive, 0
End Sub

Private Sub txtQtdePav_KeyPress(KeyAscii As Integer)
Tweak txtQtdePav, KeyAscii, IntegerPositive
End Sub
