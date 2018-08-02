VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmBaixaBancariaNovo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de de pagamento dos documentos através do banco"
   ClientHeight    =   6510
   ClientLeft      =   6375
   ClientTop       =   4350
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   12345
   Begin prjChameleon.chameleonButton btOption 
      Height          =   360
      Left            =   11160
      TabIndex        =   8
      ToolTipText     =   "Selecione uma operação"
      Top             =   6030
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   635
      BTYPE           =   14
      TX              =   "&Opções"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmBaixaBancariaNovo.frx":0000
      PICN            =   "frmBaixaBancariaNovo.frx":001C
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdParc 
      Height          =   1275
      Left            =   30
      TabIndex        =   7
      Top             =   5220
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   2249
      _Version        =   393216
      Rows            =   8
      Cols            =   17
      FixedCols       =   0
      BackColor       =   15658734
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   192
      ForeColorSel    =   16777215
      BackColorBkg    =   15658734
      GridColor       =   8421504
      GridColorFixed  =   14737632
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmBaixaBancariaNovo.frx":0107
   End
   Begin Tributacao.jcFrames frProgress 
      Height          =   465
      Left            =   4080
      Top             =   3780
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   820
      FrameColor      =   255
      FillColor       =   4210688
      TextBoxColor    =   8454016
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
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
      ThemeColor      =   3
      ColorFrom       =   192
      ColorTo         =   8438015
      Begin Tributacao.XP_ProgressBar pBar 
         Height          =   165
         Left            =   150
         TabIndex        =   6
         Top             =   180
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   291
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   255
         Scrolling       =   1
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   555
      Left            =   30
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   979
      Caption         =   ""
      TextBoxHeight   =   30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSComCtl2.DTPicker dtDataCredito 
         Height          =   315
         Left            =   1650
         TabIndex        =   10
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   82444289
         CurrentDate     =   43017
      End
      Begin VB.ComboBox cmbArquivo 
         Height          =   315
         Left            =   8220
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   2415
      End
      Begin VB.ComboBox cmbBanco 
         Height          =   315
         ItemData        =   "frmBaixaBancariaNovo.frx":01B7
         Left            =   4410
         List            =   "frmBaixaBancariaNovo.frx":01B9
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   2535
      End
      Begin prjChameleon.chameleonButton btLoad 
         Height          =   360
         Left            =   10980
         TabIndex        =   5
         ToolTipText     =   "Carregar Registro(s)"
         Top             =   105
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   635
         BTYPE           =   14
         TX              =   "&Carregar"
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmBaixaBancariaNovo.frx":01BB
         PICN            =   "frmBaixaBancariaNovo.frx":01D7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Crédito..:"
         Height          =   225
         Index           =   1
         Left            =   270
         TabIndex        =   9
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivo..:"
         Height          =   195
         Left            =   7440
         TabIndex        =   3
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco..:"
         Height          =   165
         Index           =   0
         Left            =   3750
         TabIndex        =   1
         Top             =   180
         Width           =   705
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4635
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Documento"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt.Pagam."
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Vencto"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Vl.Pago"
         Object.Width           =   1766
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "SN"
         Object.Width           =   740
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Mês"
         Object.Width           =   847
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Ano"
         Object.Width           =   1006
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "DA"
         Object.Width           =   739
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Situação retorno"
         Object.Width           =   3704
      EndProperty
   End
End
Attribute VB_Name = "frmBaixaBancariaNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Registro_Dia
    Codigo_Banco As Integer
    Nome_Banco As String
    Arquivo As String
    Valor_Total As Double
    Reg_Baixado As Integer
    Reg_Aberto As Integer
End Type


Private Type Registro
    nNumDoc As Long
    nSeq As Integer
    sDataDoc As String
    sDataPag As Date
    sDataCred As Date
    nValorPago As Double
    sAgencia As String
    nValorTarifa As Double
    sSitRetorno As String
    bExiste As Boolean
    bIsentoMJ As Boolean
    sCNPJ As String
    nAno As Integer
    nMes As Integer
    sDataVencto As Date
    nValorTarifaBancaria As Double
    nSomaTributo As Double
    sDataPagCalc As Date
    sConta As String
End Type

Private Type Documento
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    sDataVencto As String
    sSit As String
    nNumeroLivro As Integer
    nPaginaLivro As Integer
    bAjuizado As Boolean
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
    nValorTarifa As Double
    nValorDif As Double
    nValorCompensado As Double
    sBx As String
    sDp As String
    nSeqReg As Integer
    bExiste As Boolean
    sCNPJ As String
End Type

Private Type TRIBUTO
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTrib As Integer
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
    nValorTarifa As Double
    nValorCompensado As Double
    sAj As String
    sDA As String
    nFicha As Long
    nFichaJM As Long
    nFichaC As Long
End Type

Private Type TributoFicha
    nCodTrib As Integer
    sAbrevTrib As String
    Ficha As Long
    FichaJrMulta As Long
    FichaDivida As Long
    FichaDaJrMul As Long
    FichaDaEnca As Long
    FichaAjuiza As Long
    FichaAjJrMul As Long
    FichaAjEnca As Long
End Type

Private Type Ficha
    Ficha As Long
    Natureza As String
    Desc As String
    Vinculo As String
    Perc As Double
End Type

Private Type TributoProp
    nCodTrib As Integer
    nValorTrib As Double
    nPerc As Double
    nNovoValor As Double
End Type

Dim aRegistro() As Registro, aDoc() As Documento, aTrib() As TRIBUTO, aTribF() As TributoFicha, aFicha() As Ficha, sSeqArq As String
Dim aReg() As Registro_Dia, bExec As Boolean

Private Sub btLoad_Click()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, itmX As ListItem, RdoAux2 As rdoResultset, nNumDoc As Long, y As Integer, bFind As Boolean
Dim nValorEfetivo As Double, sArquivo As String

ReDim aRegistro(0)
nValorEfetivo = 0

frProgress.Visible = True
pBar.value = 0
lvMain.ListItems.Clear
nPos = 1

Sql = "select * from importacao_banco where data_credito='" & Format(dtDataCredito.value, "mm/dd/yyyy") & "'"
Sql = Sql & " and nome_arquivo='" & cmbArquivo.Text & "' and codigo_banco=" & cmbBanco.ItemData(cmbBanco.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Me.Refresh
    LockWindowUpdate lvMain.hwnd
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Set itmX = lvMain.ListItems.Add(, "C" & CStr(nPos), Format(!numero_documento, "00000000"))
        itmX.SubItems(1) = Format(!data_pagamento, "dd/mm/yyyy")
        itmX.SubItems(2) = Format(!Data_Vencimento, "dd/mm/yyyy")
        itmX.SubItems(3) = FormatNumber(!valor_pago, 2)
        itmX.SubItems(4) = IIf(!simples_nacional, "S", "N")
        itmX.SubItems(5) = IIf(IsNull(!Mes), "---", Format(!Mes, "00"))
        itmX.SubItems(6) = IIf(IsNull(!Ano), "---", !Ano)
        itmX.SubItems(7) = IIf(!debito_automatico, "S", "N")
        itmX.SubItems(8) = SubNull(!situacao_retorno)
                        
        nNumDoc = !numero_documento
        Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        ReDim Preserve aRegistro(UBound(aRegistro) + 1)
        If RdoAux2.RowCount > 0 Then
            aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
            aRegistro(UBound(aRegistro)).nSeq = nPos
            aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux2!DATADOCUMENTO, "dd/mm/yyyy")
            aRegistro(UBound(aRegistro)).sDataPag = Format(!data_pagamento, "dd/mm/yyyy")
            aRegistro(UBound(aRegistro)).sDataCred = Format(dtDataCredito.value, "dd/mm/yyyy")
            aRegistro(UBound(aRegistro)).nValorPago = FormatNumber(!valor_pago, 2)
            aRegistro(UBound(aRegistro)).nValorTarifaBancaria = 0
            aRegistro(UBound(aRegistro)).sAgencia = SubNull(!Agencia)
            aRegistro(UBound(aRegistro)).nValorTarifa = 0
            aRegistro(UBound(aRegistro)).sSitRetorno = SubNull(!situacao_retorno)
            aRegistro(UBound(aRegistro)).bExiste = True
            aRegistro(UBound(aRegistro)).bIsentoMJ = IIf(Val(SubNull(RdoAux2!isentomj)) = 0, False, True)
            aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(!data_pagamento)
        Else
            aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
            aRegistro(UBound(aRegistro)).nSeq = nPos
            aRegistro(UBound(aRegistro)).sDataDoc = ""
            aRegistro(UBound(aRegistro)).sDataPag = Format(!data_pagamento, "dd/mm/yyyy")
            aRegistro(UBound(aRegistro)).sDataCred = Format(dtDataCredito.value, "dd/mm/yyyy")
            aRegistro(UBound(aRegistro)).nValorPago = FormatNumber(!valor_pago, 2)
            aRegistro(UBound(aRegistro)).nValorTarifaBancaria = 0
            aRegistro(UBound(aRegistro)).sAgencia = SubNull(!Agencia)
            aRegistro(UBound(aRegistro)).nValorTarifa = 0
            aRegistro(UBound(aRegistro)).sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
            aRegistro(UBound(aRegistro)).bExiste = False
            aRegistro(UBound(aRegistro)).bIsentoMJ = False
            aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(!data_pagamento)
        End If
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
LockWindowUpdate 0&
pBar.value = 0
frProgress.Visible = False

End Sub

Private Sub cmbArquivo_Click()
Dim x As Integer, sBanco As String
Exit Sub
If Not bExec Then Exit Sub
If cmbBanco.ListIndex > 0 Then Exit Sub
If cmbArquivo.ListIndex < 1 Then
    cmbBanco.ListIndex = 0
    Exit Sub
End If

For x = 1 To grdMain.Rows - 1
    If grdMain.TextMatrix(x, 1) = cmbArquivo.Text Then
        sBanco = grdMain.TextMatrix(x, 0)
        Exit For
    End If
Next

bExec = False
For x = 0 To cmbBanco.ListCount
    If cmbBanco.List(x) = sBanco Then
        cmbBanco.ListIndex = x
        Exit For
    End If
Next
bExec = True

End Sub

Private Sub cmbBanco_Click()
Dim nCodBanco As Integer, bFind As Boolean, x As Integer, RdoAux As rdoResultset, Sql As String

If Not bExec Then Exit Sub
If cmbBanco.ListIndex = -1 Then Exit Sub

cmbArquivo.Clear
nCodBanco = cmbBanco.ItemData(cmbBanco.ListIndex)

Sql = "select distinct nome_arquivo from importacao_banco where data_credito='" & Format(dtDataCredito.value, "mm/dd/yyyy") & "' and codigo_banco=" & cmbBanco.ItemData(cmbBanco.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbArquivo.AddItem !nome_arquivo
       .MoveNext
    Loop
   .Close
End With
cmbArquivo.ListIndex = 0

End Sub

Private Sub dtDataCredito_Change()
Dim Sql As String, RdoAux As rdoResultset, x As Integer, y As Integer, bFind As Boolean
Dim nPosReg As Integer

cmbBanco.Clear
cmbArquivo.Clear
ReDim aReg(0)
lvMain.ListItems.Clear
Sql = "SELECT * FROM importacao_banco LEFT OUTER JOIN "
Sql = Sql & "banco ON importacao_banco.Codigo_Banco = banco.codbanco where data_credito='" & Format(dtDataCredito.value, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aReg(UBound(aReg) + 1)
        nPosReg = UBound(aReg)
        aReg(nPosReg).Codigo_Banco = !Codigo_Banco
        aReg(nPosReg).Nome_Banco = IIf(IsNull(!NomeBanco), "Outro banco", !NomeBanco)
        aReg(nPosReg).Arquivo = GetFileNameFromPath(!nome_arquivo)
        aReg(nPosReg).Valor_Total = !valor_pago
        If IsNull(!data_controle) Then
            aReg(nPosReg).Reg_Baixado = 0
        Else
            aReg(nPosReg).Reg_Baixado = 1
        End If
       .MoveNext
    Loop
   .Close
End With

cmbBanco.Clear
For x = 1 To UBound(aReg)
    With aReg(x)
        
        bFind = False
        For y = 0 To cmbBanco.ListCount - 1
            If cmbBanco.ItemData(y) = .Codigo_Banco Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
            cmbBanco.AddItem .Nome_Banco
            cmbBanco.ItemData(cmbBanco.NewIndex) = .Codigo_Banco
        End If
        
    End With
Next

If cmbBanco.ListCount > 0 Then
    cmbBanco.ListIndex = 0
End If
cmbBanco_Click
bExec = True


End Sub

Private Sub Form_Load()
Centraliza Me
grdParc.Rows = 1
End Sub


Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvMain.SortKey = ColumnHeader.Position - 1
lvMain.Sorted = True
lvMain.SortOrder = lvwAscending

End Sub


Private Sub CallPb(nPos As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents
If nTotal = 0 Then Exit Sub
If ((nPos * 100) / nTotal) <= 100 Then
   pBar.value = (nPos * 100) / nTotal
Else
   pBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Private Sub CarregaParcela(nNumDoc As Long, nLinha As Integer)

End Sub
