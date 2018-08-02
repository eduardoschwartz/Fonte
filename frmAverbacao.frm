VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmAverbacao 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certidão de Averbação"
   ClientHeight    =   5100
   ClientLeft      =   3645
   ClientTop       =   2415
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6585
   Begin VB.TextBox txtResp 
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   6
      Top             =   120
      Width           =   4905
   End
   Begin VB.TextBox txtCod 
      Height          =   315
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   5
      Top             =   480
      Width           =   1275
   End
   Begin VB.ComboBox cmbBairro 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   4515
   End
   Begin prjChameleon.chameleonButton cmdClear 
      Height          =   345
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Limpar Lista"
      Top             =   4680
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Limpar"
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
      MICON           =   "frmAverbacao.frx":0000
      PICN            =   "frmAverbacao.frx":001C
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
      Left            =   5220
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAverbacao.frx":0176
      PICN            =   "frmAverbacao.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3780
      TabIndex        =   0
      ToolTipText     =   "Imprime as Certidões"
      Top             =   4680
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   609
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
      MICON           =   "frmAverbacao.frx":0200
      PICN            =   "frmAverbacao.frx":021C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1770
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmAverbacao.frx":0376
      Top             =   4110
      Visible         =   0   'False
      Width           =   1395
   End
   Begin prjChameleon.chameleonButton cmdAddCod 
      Height          =   285
      Left            =   2910
      TabIndex        =   7
      ToolTipText     =   "Atualizar Lista"
      Top             =   480
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAverbacao.frx":037C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAddBairro 
      Height          =   285
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Atualizar Lista"
      Top             =   840
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAverbacao.frx":0398
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdMain 
      Height          =   3315
      Left            =   30
      TabIndex        =   12
      Top             =   1260
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5847
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderButtons   =   0   'False
      HeaderDragReorderColumns=   0   'False
      HeaderHotTrack  =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Responsável........:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido..:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   540
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro Inteiro.........:"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   1425
   End
End
Attribute VB_Name = "frmAverbacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset
Dim xImovel As clsImovel, nVaVT As Double, nVaVP As Double, nAnoCalculo As Integer
'TIPOS
Private Type PROFUNDIDADE
    Distrito As Integer
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORPROFUN
    Distrito As Integer
    Codigo As Integer
    Fator As Double
End Type
Private Type GLEBA
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORCATEG
    Uso As Integer
    Tipo As Integer
    Categoria As Integer
    Fator As Double
End Type
'MATRIZES
Dim aFatorD() As Double
Dim aFatorP() As Double
Dim aFatorT() As Double
Dim aFatorS() As Double
Dim aFatorG() As Double
Dim aFatorR() As Double
Dim aProf() As PROFUNDIDADE
Dim aFatorF() As FATORPROFUN
Dim aFatorC() As FATORCATEG
Dim aGleba() As GLEBA


Private Sub cmdAddBairro_Click()

If cmbBairro.ListIndex = -1 Then
    MsgBox "Selecione um bairro.", vbExclamation, "Atenção"
    Exit Sub
End If

grdMain.Redraw = False
grdMain.Clear


Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE LI_CODBAIRRO=" & cmbBairro.ItemData(cmbBairro.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, Format(!CODREDUZIDO, "000000"), DT_CENTER
        
       .MoveNext
    Loop
   .Close
End With
grdMain.Redraw = True

End Sub

Private Sub cmdAddCod_Click()
Dim X As Integer
Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Imóvel não localizado.", vbExclamation, "Atenção"
        Exit Sub
    End If
   .Close
End With

For X = 1 To grdMain.Rows
    If Val(grdMain.cell(X, 1).Text) = Val(txtCod.Text) Then
        MsgBox "Imóvel já incluido.", vbExclamation, "Atenção"
        Exit Sub
    End If
Next

grdMain.AddRow
grdMain.CellDetails grdMain.Rows, 1, Format(txtCod.Text, "000000"), DT_CENTER
txtCod.Text = ""
End Sub

Private Sub cmdClear_Click()
grdMain.Clear
End Sub

Private Sub cmdPrint_Click()
Dim bAchou As Boolean

If Trim$(txtResp.Text) = "" Then
    MsgBox "Digite o nome do responsavel.", vbExclamation, "Atenção"
    Exit Sub
End If

If grdMain.Rows = 0 Then
    MsgBox "Não existem imóveis selecionados.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False
With grdMain
    For X = 1 To .Rows
        If Trim$(.cell(X, 2).Text) <> "" Or Trim$(.cell(X, 3).Text) <> "" Then
           bAchou = True
           Exit For
        End If
    Next
End With

If Not bAchou Then
    MsgBox "Nenhum imóveis foi selecionado.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False
With grdMain
    For X = 1 To .Rows
        If Trim$(.cell(X, 2).Text) = "" And Trim$(.cell(X, 3).Text) = "" Then
        Else
            If Trim$(.cell(X, 2).Text) = "" Or Trim$(.cell(X, 3).Text) = "" Then
                bAchou = True
                Exit For
            End If
        End If
    Next
End With

If bAchou Then
    MsgBox "Alguns imóveis estão com dados incompletos.", vbExclamation, "Atenção"
    Exit Sub
End If

Grava
frmReport.ShowReport "AVERBACAO", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

nAnoCalculo = Year(Now)
Centraliza Me
GridHeader
CarregaBairro
Set xImovel = New clsImovel
LoadMatrix
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub GridHeader()
With grdMain
    .GridFillLineColor = vbWhite
    .Editable = True
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
        
    .AddColumn "kCod", "Código", ecgHdrTextALignCentre, , 60
    .AddColumn "kCert", "Certidão", ecgHdrTextALignLeft, , 70
    .AddColumn "kProc", "Processo", ecgHdrTextALignLeft, , 90
    .AddColumn "kReq", "Requerente", ecgHdrTextALignLeft, , 190
End With

End Sub

Private Sub CarregaBairro()

Sql = "Select CODBAIRRO,DESCBAIRRO From BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO<>999 "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       cmbBairro.AddItem !DescBairro
       cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub grdMain_CancelEdit()
txtEdit.Visible = False
End Sub

Private Sub grdMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim bSelection As Boolean
   bSelection = ((grdMain.SelectedRow > 0) And (grdMain.SelectedCol > 0))
   If (bSelection) Then
         ' Check the cell boundary:
         Dim lLeft As Long
         Dim lTop As Long
         Dim lWidth As Long
         Dim lHeight As Long
         grdMain.CellBoundary grdMain.SelectedRow, grdMain.SelectedCol, lLeft, lTop, lWidth, lHeight
         'Debug.Print lLeft, lTop, lWidth, lHeight
   End If

End Sub

Private Sub grdMain_PreCancelEdit(ByVal lrow As Long, ByVal lcol As Long, newvalue As Variant, bStayInEditMode As Boolean)
Dim nCodCidadao As Long
If lcol <> 2 Then
    If txtEdit.Text = "" Then
        grdMain.CellText(grdMain.EditRow, 3) = ""
        grdMain.CellText(grdMain.EditRow, 4) = ""
        Exit Sub
    End If
    sValidaProc = ValidaProcesso(txtEdit.Text)
    'If sValidaProc <> "OK" Then
    If InStr(1, sValidaProc, "ARQUIVADO", vbBinaryCompare) = 0 And InStr(1, sValidaProc, "CANCELADO", vbBinaryCompare) = 0 And sValidaProc <> "OK" Then
        MsgBox sValidaProc, vbCritical, "Atenção"
        txtEdit.Text = ""
        bStayInEditMode = True
        Exit Sub
    Else
        Sql = "SELECT CODCIDADAO FROM PROCESSOGTI WHERE ANO=" & ExtraiAnoProcesso(txtEdit.Text) & " AND NUMERO=" & ExtraiNumeroProcesso(txtEdit.Text)
        'Sql = "SELECT CODCIDADAO FROM PROCESSOGTI WHERE ANO=" & Val(Right$(txtEdit.Text, 4)) & " AND NUMERO=" & Val(Left$(txtEdit.Text, Len(txtEdit.Text) - 5))
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Cidadão não localizado no protocolo.", vbExclamation, "Atenção"
                Exit Sub
            Else
                nCodCidadao = !CodCidadao
            End If
           .Close
        End With
        Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                grdMain.CellText(grdMain.EditRow, 4) = !nomecidadao
            Else
                grdMain.CellText(grdMain.EditRow, 4) = "PREFEITURA MUNICIPAL DE JABOTICABAL"
            End If
           .Close
        End With
        grdMain.CellText(grdMain.EditRow, grdMain.EditCol) = txtEdit.Text
    End If
Else
    grdMain.CellText(grdMain.EditRow, grdMain.EditCol) = txtEdit.Text
End If
End Sub

Private Sub grdMain_RequestEdit(ByVal lrow As Long, ByVal lcol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
Dim sText As String
   
   
   ' Don't allow editing the icon-only columns:
   If (grdMain.ColumnKey(lcol) <> "kProc") And (grdMain.ColumnKey(lcol) <> "kCert") Then
      bCancel = True
      
      Exit Sub
   End If
   
   ' Get boundary of the cell:
   grdMain.CellBoundary lrow, lcol, lLeft, lTop, lWidth, lHeight
   
   ' Get the text:
   If Not IsMissing(grdMain.CellText(lrow, lcol)) Then
      sText = grdMain.CellFormattedText(lrow, lcol)
   Else
      sText = ""
   End If
   
   ' If the user has initiated edit mode by a key, we want
   ' to add this to the text.  This is really a common
   ' thing and should probably be supported automatically
   ' in the grid:
   If Not (iKeyAscii = 0) Then
      sText = Chr$(iKeyAscii) & sText
      txtEdit.Text = sText
      txtEdit.SelStart = 1
      txtEdit.SelLength = Len(sText)
   Else
      txtEdit.Text = sText
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(sText)
   End If
   
   ' Set the text properties to match the grid cell being edited:
   Set txtEdit.Font = grdMain.CellFont(lrow, lcol)
   If grdMain.CellBackColor(lrow, lcol) = -1 Then
      txtEdit.BackColor = grdMain.BackColor
   Else
      txtEdit.BackColor = grdMain.CellBackColor(lrow, lcol)
   End If
   
   ' Move the text box to the edit position, make it visible and give it the focus:
   txtEdit.Move lLeft + grdMain.Left, lTop + grdMain.Top - 30 + Screen.TwipsPerPixelY, lWidth, lHeight
   txtEdit.Visible = True
   txtEdit.ZOrder
   txtEdit.SetFocus

End Sub


Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode = vbKeyReturn) Then
      ' Request Commit edit.  This will fire the
      ' grid's PreCancelEdit event, which gives you
      ' an opportunity to validate the data and put
      ' it in the cell if good.  The CancelEdit
      ' event will then fire afterwards.
      grdMain.EndEdit
   ElseIf (KeyCode = vbKeyEscape) Then
      ' Cancel edit.  This skips PreCancelEdit and
      ' fires the CancelEdit event
      grdMain.CancelEdit
   ElseIf (grdMain.SingleClickEdit) Then
      Select Case KeyCode
              
      End Select
   End If

End Sub

Private Sub Grava()
Dim sInsc As String, sBairro As String, sProp As String, nVVP As Double, nVVT As Double, sResp As String
Dim sVVPE As String, sVVTE As String

Sql = "DELETE FROM AVERBACAO WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

sResp = UCase$(Trim$(txtResp.Text))

With grdMain
    For X = 1 To .Rows
        If Trim$(.cell(X, 2).Text) <> "" And Trim$(.cell(X, 3).Text) <> "" Then
            xImovel.CarregaImovel Val(.cell(X, 1).Text)
            sBairro = xImovel.DescBairro
            sProp = xImovel.NomePropPrincipal
            sInsc = xImovel.Inscricao
            Calculo Val(.cell(X, 1).Text)
            nVVT = nVaVT
            nVVP = nVaVP
            sVVTE = Extenso(nVVT)
            sVVPE = Extenso(nVVP)
            Sql = "INSERT AVERBACAO(COMPUTER,CODREDUZIDO,PROCESSO,INSCRICAO,REQUERENTE,BAIRRO,PROPRIETARIO,VVP,"
            Sql = Sql & "VVT,RESPONSAVEL,VVPE,VVTE,CERTIDAO) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & Val(.cell(X, 1).Text) & ",'" & .cell(X, 3).Text & "','" & sInsc & "','"
            Sql = Sql & Mask(.cell(X, 4).Text) & "','" & sBairro & "','" & Mask(sProp) & "'," & Virg2Ponto(CStr(nVVP)) & "," & Virg2Ponto(CStr(nVVT)) & ",'" & sResp & "','"
            Sql = Sql & sVVPE & "','" & sVVTE & "','" & .cell(X, 2).Text & "')"
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With

End Sub

Private Sub Calculo(nCodReduz)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, X As Integer


nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where (CADIMOB.CODREDUZIDO = " & nCodReduz & ") GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = !Li_CodBairro
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nTestadaPrincipal = 1
            nTestada1 = 1
        Else
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                nTestadaPrincipal = nSomaTestada / nNumTestadas
            End If
        End If
       .Close
    End With
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    'BUSCA ÁREA PRINCIPAL
    'Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 1 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) And RdoAux2!QTDEPAV < 2 And nAreaTerreno < 600 Then
                        If nAnoCalculo > 2006 Then
                            Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
                            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                            If RdoAux4.RowCount = 0 Then
                                bIsento = True
                            End If
                            RdoAux4.Close
                        Else
                            bIsento = True
                        End If
                    End If
                End If
               .Close
            End With
        Else
            bIsento = False
        End If
        
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
        
        'novo VVP ***********************************
        If nAnoCalculo > 2007 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            If bTemPredial Then
                Do Until .EOF
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                    nFatorCategoria = 0
                    For X = 1 To UBound(aFatorC)
                        If aFatorC(X).Uso = nUso And aFatorC(X).Tipo = nTipo And aFatorC(X).Categoria = nCat Then
                           nFatorCategoria = aFatorC(X).Fator
                           Exit For
                        End If
                    Next
                    nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
                   .MoveNext
                Loop
            End If
        Else
            If bTemPredial Then
                 nUso = !USOCONSTR
                 nTipo = !TIPOCONSTR
                 nCat = !CATCONSTR
            End If
        End If
       .Close
    End With
    
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    For X = 1 To UBound(aGleba)
        If nAreaTerreno >= aGleba(X).Min And nAreaTerreno <= aGleba(X).Max Then
             Exit For
        ElseIf nAreaTerreno >= aGleba(X).Min And aGleba(X).Max = 0 Then
             Exit For
        End If
    Next
    nCodGleba = aGleba(X).Codigo
    'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
    nFatorGleba = aFatorG(nCodGleba)
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For X = 1 To UBound(aProf)
            If aProf(X).Distrito = !Distrito Then
               If nValorProfundidade >= CDbl(FormatNumber(aProf(X).Min, 2)) And nValorProfundidade <= CDbl(FormatNumber(aProf(X).Max, 2)) Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(X).Min And aProf(X).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(X).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For X = 1 To UBound(aFatorF)
            If aFatorF(X).Distrito = !Distrito And aFatorF(X).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(X).Fator
               Exit For
            End If
        Next
     Else
        nFatorProfundidade = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba, 2)
    'CÁLCULO VALOR VENAL TERRITORIAL
    nFatorDistrito = aFatorD(!Distrito)
    nValorFatores = nValorFatores * nFatorDistrito
    nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        If nAnoCalculo < 2008 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            For X = 1 To UBound(aFatorC)
                If aFatorC(X).Uso = nUso And aFatorC(X).Tipo = nTipo And aFatorC(X).Categoria = nCat Then
                   nFatorCategoria = aFatorC(X).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
        End If
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
    Else
        nFatorDistrito = 0
        nFatorCategoria = 0
    End If
    
    
    'VALOR ITU/IPTU
    nVaVP = nValorVenalPredial
    nVaVT = nValorVenalTerritorial
End With

End Sub

Private Sub CalculoOld(nCodReduz)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, X As Integer


nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nNumTestadas = 1
            nTestada1 = 1
            nTestadaPrincipal = 1
            GoTo 2
        End If
        If nNumTestadas = 1 Then
            nTestadaPrincipal = !AREATESTADA
            nTestada1 = !AREATESTADA
        Else
            nSomaTestada = 0
            Do Until .EOF
               If !NUMFACE = RdoAux!Seq Then
                  nTestada1 = !AREATESTADA
               End If
               nSomaTestada = nSomaTestada + !AREATESTADA
              .MoveNext
            Loop
            nTestadaPrincipal = nSomaTestada / nNumTestadas
        End If
    End With
2:
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    'BUSCA ÁREA PRINCIPAL
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If Not IsNull(!soma) Then
                If !soma <= 65 And RdoAux2!USOCONSTR = 0 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) Then
                    bIsento = True
                End If
            End If
           .Close
        End With
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
        If bTemPredial Then
             nUso = !USOCONSTR
             nTipo = !TIPOCONSTR
             nCat = !CATCONSTR
             Select Case !USOCONSTR
                  Case 0
                     nTaxaLimpeza = 3.78
                  Case 1, 2, 3, 4, 5
                     nTaxaLimpeza = 10.57
                  Case Else
                     nTaxaLimpeza = 3.01
             End Select
        Else
             nTaxaLimpeza = 3.01
        End If
        nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
       '--CÁLCULO DA TAXA DE CONSERVAÇÃO
        If RdoAux!PAVIMENTO = 1 Then
           nTaxaConservacao = 1.35 * nTestadaPrincipal
        Else
           nTaxaConservacao = 0
        End If
        If nCodBairro = 81 Then
           nTaxaLimpeza = 1
           nTaxaConservacao = 1
        End If
       .Close
    End With
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    'If !Dt_CodUsoTerreno = 6 Then
        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
        For X = 1 To UBound(aGleba)
            If nAreaTerreno >= aGleba(X).Min And nAreaTerreno <= aGleba(X).Max Then
                 Exit For
            ElseIf nAreaTerreno >= aGleba(X).Min And aGleba(X).Max = 0 Then
                 Exit For
            End If
        Next
        nCodGleba = aGleba(X).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
        nFatorGleba = aFatorG(nCodGleba)
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
    'Else
    '    nFatorGleba = 1
    'End If
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For X = 1 To UBound(aProf)
            If aProf(X).Distrito = !Distrito Then
               If nValorProfundidade >= Round(aProf(X).Min, 2) And nValorProfundidade <= Round(aProf(X).Max, 2) Then
                  Exit For
               ElseIf nValorProfundidade >= Round(aProf(X).Min, 2) And Round(aProf(X).Max, 2) = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(X).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For X = 1 To UBound(aFatorF)
            If aFatorF(X).Distrito = !Distrito And aFatorF(X).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(X).Fator
               Exit For
            End If
        Next
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
     Else
        nFatorProfundidade = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    'FATOR SITUACAO 98
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    'FATOR PEDOLOGIA 98
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    'FATOR TOPOGRAFIA 98
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
    'CÁLCULO VALOR VENAL TERRITORIAL
    nValorVenalTerritorial = Round(nAreaTerreno, 2) * Round(nValorAgrupamento, 2) * Round(nValorFatores, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba * nFatorDistrito
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = Round(nAreaTerreno, 2) * Round(nValorAgrupamento, 2) * Round(nValorFatores, 2)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        nValorVenalPredial = 0
        For X = 1 To UBound(aFatorC)
            If aFatorC(X).Uso = nUso And aFatorC(X).Tipo = nTipo And aFatorC(X).Categoria = nCat Then
               nFatorCategoria = aFatorC(X).Fator
               Exit For
            End If
        Next
        nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
        
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
    Else
        nValorVenalPredial = 0
    End If
    'VALOR ITU/IPTU
    nVaVP = nValorVenalPredial
    nVaVT = nValorVenalTerritorial
End With

End Sub

Private Sub LoadMatrix()

ReDim aFatorD(3)
ReDim aFatorP(6)
ReDim aFatorT(6)
ReDim aFatorS(6)
ReDim aFatorG(23)
ReDim aFatorR(7)

Sql = "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & nAnoCalculo & " ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE CODAGRUPAMENTO<8 AND ANOFATOR=" & nAnoCalculo & " AND  CODMOEDA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        aFatorP(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .Close
End With

ReDim aProf(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN FROM PROFUNDIDADE ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aProf(UBound(aProf) + 1)
        aProf(UBound(aProf)).Distrito = !CODDISTRITO
        aProf(UBound(aProf)).Codigo = !CODPROFUN
        aProf(UBound(aProf)).Min = !MINPROFUN
        aProf(UBound(aProf)).Max = !MAXPROFUN
       .MoveNext
     Loop
    .Close
End With


ReDim aFatorF(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN; " & _
      "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN= 1998 ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorF(UBound(aFatorF) + 1)
        aFatorF(UBound(aFatorF)).Distrito = !CODDISTRITO
        aFatorF(UBound(aFatorF)).Codigo = !CODPROFUN
        aFatorF(UBound(aFatorF)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .Close
End With

ReDim aGleba(0)
Sql = "SELECT CODGLEBA,MINGLEBA,MAXGLEBA FROM GLEBA ORDER BY CODGLEBA "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aGleba(UBound(aGleba) + 1)
        aGleba(UBound(aGleba)).Codigo = !CODGLEBA
        aGleba(UBound(aGleba)).Min = !MINGLEBA
        aGleba(UBound(aGleba)).Max = !MAXGLEBA
       .MoveNext
     Loop
    .Close
End With

ReDim aFatorC(0)
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1; " & _
      "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=1998 AND CODMOEDA=1 "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorC(UBound(aFatorC) + 1)
        aFatorC(UBound(aFatorC)).Uso = !CODUSO
        aFatorC(UBound(aFatorC)).Tipo = !CodTipo
        aFatorC(UBound(aFatorC)).Categoria = !CODCATEG
        aFatorC(UBound(aFatorC)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub


