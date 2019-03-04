VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmBaixaBancariaNovo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa Bancária Integrada"
   ClientHeight    =   7635
   ClientLeft      =   5595
   ClientTop       =   4995
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   14280
   Begin MSComctlLib.ListView lvMain 
      Height          =   4785
      Left            =   30
      TabIndex        =   9
      Top             =   2430
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   8440
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Banco"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Arquivo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Pagam."
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Dt.Vencto"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Vl.Pago"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "SN"
         Object.Width           =   705
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Mês"
         Object.Width           =   847
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Ano"
         Object.Width           =   971
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DA"
         Object.Width           =   704
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Retorno do Banco"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Ag."
         Object.Width           =   1306
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "C/C"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Situação"
         Object.Width           =   2717
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   7320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   2415
      Left            =   0
      Top             =   0
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   4260
      FrameColor      =   8421504
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
      ColorFrom       =   0
      ColorTo         =   0
      Begin prjChameleon.chameleonButton btOption 
         Height          =   375
         Left            =   12540
         TabIndex        =   7
         ToolTipText     =   "Selecione uma operação"
         Top             =   1830
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
         BTYPE           =   3
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
         COLTYPE         =   2
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
      Begin prjChameleon.chameleonButton btLoad 
         Height          =   375
         Left            =   11070
         TabIndex        =   6
         ToolTipText     =   "Carregar Registro(s)"
         Top             =   1830
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
         BTYPE           =   3
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmBaixaBancariaNovo.frx":0107
         PICN            =   "frmBaixaBancariaNovo.frx":0123
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cmbArquivo 
         Height          =   315
         Left            =   10860
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3285
      End
      Begin VB.ComboBox cmbBanco 
         Height          =   315
         Left            =   10860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   390
         Width           =   3285
      End
      Begin MSFlexGridLib.MSFlexGrid grdMain 
         Height          =   2295
         Left            =   2550
         TabIndex        =   1
         Top             =   45
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   8421376
         ForeColorFixed  =   16777215
         BackColorSel    =   255
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         MergeCells      =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "<Nome do Banco                 |<Nome dos Arquivos     |>Valor Total        |>Valor Baixado          |>Valor Aberto          "
      End
      Begin MSComCtl2.MonthView mvData 
         Height          =   2310
         Left            =   60
         TabIndex        =   0
         Top             =   30
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "frmBaixaBancariaNovo.frx":04C8
         MonthBackColor  =   16777215
         StartOfWeek     =   125894657
         TitleBackColor  =   8421376
         TitleForeColor  =   16777215
         TrailingForeColor=   16711935
         CurrentDate     =   42844
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivo"
         Height          =   195
         Left            =   10890
         TabIndex        =   4
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   165
         Left            =   10860
         TabIndex        =   2
         Top             =   150
         Width           =   705
      End
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
    Valor_Baixado As Double
    Reg_Baixado As Integer
    Valor_Aberto As Double
    Reg_Aberto As Integer
End Type

Private Type Resumo_Dia
    Codigo_Banco As Integer
    Nome_Banco As String
    Arquivo As String
    Valor_Total As Double
    Valor_Baixado As Double
    Reg_Baixado As Integer
    Valor_Aberto As Double
    Reg_Aberto As Integer
End Type

Dim aReg() As Registro_Dia, aResumo() As Resumo_Dia, bExec As Boolean

Private Sub btLoad_Click()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, itmX As ListItem

PBar.value = 0
lvMain.ListItems.Clear
nPos = 1

Sql = "select * from importacao_banco where data_credito='" & Format(mvData.value, "mm/dd/yyyy") & "'"
If cmbArquivo.ListIndex > 0 Then
    Sql = Sql & " and nome_arquivo='" & cmbArquivo.Text & "'"
End If
If cmbBanco.ListIndex > 0 Then
    Sql = Sql & " and codigo_banco=" & cmbBanco.ItemData(cmbBanco.ListIndex)
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Set itmX = lvMain.ListItems.Add(, "C" & CStr(nPos), Format(!Codigo_Banco, "000"))
        itmX.SubItems(1) = GetFileNameFromPath(!nome_arquivo)
        itmX.SubItems(2) = Format(!numero_documento, "00000000")
        itmX.SubItems(3) = Format(!Data_Pagamento, "dd/mm/yyyy")
        itmX.SubItems(4) = Format(!Data_Vencimento, "dd/mm/yyyy")
        itmX.SubItems(5) = FormatNumber(!valor_pago, 2)
        itmX.SubItems(6) = IIf(!simples_nacional, "S", "N")
        itmX.SubItems(7) = IIf(IsNull(!Mes), "---", Format(!Mes, "00"))
        itmX.SubItems(8) = IIf(IsNull(!Ano), "---", !Ano)
        itmX.SubItems(9) = IIf(!debito_automatico, "S", "N")
        itmX.SubItems(10) = SubNull(!situacao_retorno)
        itmX.SubItems(11) = SubNull(!Agencia)
        itmX.SubItems(12) = SubNull(!conta_corrente)
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
PBar.value = 0

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
Dim nCodBanco As Integer, bFind As Boolean, x As Integer

If Not bExec Then Exit Sub
If cmbBanco.ListIndex = -1 Then Exit Sub

cmbArquivo.Clear
cmbArquivo.AddItem "(Todos os Arquivos)"
cmbArquivo.ItemData(cmbArquivo.NewIndex) = 0
nCodBanco = cmbBanco.ItemData(cmbBanco.ListIndex)


For x = 1 To grdMain.Rows - 1
    If nCodBanco > 0 Then
        If grdMain.TextMatrix(x, 0) = cmbBanco.Text Then
            cmbArquivo.AddItem grdMain.TextMatrix(x, 1)
        End If
    Else
        cmbArquivo.AddItem grdMain.TextMatrix(x, 1)
    End If
Next
cmbArquivo.ListIndex = 0

End Sub

Private Sub Form_Load()
Centraliza Me
mvData.value = Now
mvData_DateClick (Now)
End Sub


Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvMain.SortKey = ColumnHeader.Position - 1
lvMain.Sorted = True
lvMain.SortOrder = lvwAscending

End Sub

Private Sub mvData_DateClick(ByVal DateClicked As Date)
Dim Sql As String, RdoAux As rdoResultset, x As Integer, Y As Integer, bFind As Boolean
Dim nPosReg As Integer

ReDim aReg(0): ReDim aResumo(0)
grdMain.Rows = 1

Sql = "SELECT * FROM importacao_banco LEFT OUTER JOIN "
Sql = Sql & "banco ON importacao_banco.Codigo_Banco = banco.codbanco where data_credito='" & Format(mvData.value, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aReg(UBound(aReg) + 1)
        nPosReg = UBound(aReg)
       ' If !Codigo_Banco = 104 Then MsgBox "teste"
        aReg(nPosReg).Codigo_Banco = !Codigo_Banco
        aReg(nPosReg).Nome_Banco = IIf(IsNull(!NomeBanco), "Outro banco", !NomeBanco)
        aReg(nPosReg).Arquivo = GetFileNameFromPath(!nome_arquivo)
        aReg(nPosReg).Valor_Total = !valor_pago
        If IsNull(!data_controle) Then
            aReg(nPosReg).Valor_Aberto = !valor_pago
            aReg(nPosReg).Reg_Aberto = 1
            aReg(nPosReg).Valor_Baixado = 0
            aReg(nPosReg).Reg_Baixado = 0
        Else
            aReg(nPosReg).Valor_Aberto = 0
            aReg(nPosReg).Reg_Aberto = 0
            aReg(nPosReg).Valor_Baixado = !valor_pago
            aReg(nPosReg).Reg_Baixado = 1
        End If
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(aReg)
    With aReg(x)
        
        bFind = False
        For Y = 1 To UBound(aResumo)
            If aResumo(Y).Codigo_Banco = aReg(x).Codigo_Banco And aResumo(Y).Arquivo = aReg(x).Arquivo Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
            ReDim Preserve aResumo(UBound(aResumo) + 1)
            nPosReg = UBound(aResumo)
            aResumo(nPosReg).Codigo_Banco = aReg(x).Codigo_Banco
            aResumo(nPosReg).Nome_Banco = IIf(IsNull(aReg(x).Nome_Banco), "Outro banco", aReg(x).Nome_Banco)
            aResumo(nPosReg).Arquivo = aReg(x).Arquivo
            aResumo(nPosReg).Valor_Total = aReg(x).Valor_Total
            aResumo(nPosReg).Reg_Aberto = aResumo(nPosReg).Reg_Aberto + aReg(x).Reg_Aberto
            aResumo(nPosReg).Reg_Baixado = aResumo(nPosReg).Reg_Baixado + aReg(x).Reg_Baixado
            aResumo(nPosReg).Valor_Aberto = aResumo(nPosReg).Valor_Aberto + aReg(x).Valor_Aberto
            aResumo(nPosReg).Valor_Baixado = aResumo(nPosReg).Valor_Baixado + aReg(x).Valor_Baixado
        Else
            aResumo(Y).Valor_Total = aResumo(Y).Valor_Total + aReg(x).Valor_Total
            aResumo(Y).Reg_Aberto = aResumo(Y).Reg_Aberto + aReg(x).Reg_Aberto
            aResumo(Y).Reg_Baixado = aResumo(Y).Reg_Baixado + aReg(x).Reg_Baixado
            aResumo(Y).Valor_Aberto = aResumo(Y).Valor_Aberto + aReg(x).Valor_Aberto
            aResumo(Y).Valor_Baixado = aResumo(Y).Valor_Baixado + aReg(x).Valor_Baixado
        End If
    End With
Next

For x = 1 To UBound(aResumo)
    With aResumo(x)
        grdMain.AddItem IIf(IsNull(.Nome_Banco), Format(.Codigo_Banco, "000") & "-Outro banco", Format(.Codigo_Banco, "000") & "-" & .Nome_Banco) & Chr(9) & .Arquivo & Chr(9) & FormatNumber(.Valor_Total, 2) & Chr(9) & _
        FormatNumber(.Valor_Baixado, 2) & "(" & .Reg_Baixado & ")" & Chr(9) & FormatNumber(.Valor_Aberto, 2) & "(" & .Reg_Aberto & ")"
    End With
Next

cmbBanco.Clear
cmbBanco.AddItem "(Todos os Bancos)"
cmbBanco.ItemData(cmbBanco.NewIndex) = 0
With grdMain
    .MergeCells = flexMergeRestrictColumns
    .MergeCol(0) = True: .MergeCol(1) = True
    For x = 1 To .Rows - 1
        .col = 3
        .Row = x
        .CellForeColor = &H8000&
        .col = 4
        .CellForeColor = vbRed
        bFind = False
        For Y = 0 To cmbBanco.ListCount - 1
            If cmbBanco.ItemData(Y) = Val(Left(.TextMatrix(x, 0), 3)) Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
            cmbBanco.AddItem .TextMatrix(x, 0)
            cmbBanco.ItemData(cmbBanco.NewIndex) = Val(Left(.TextMatrix(x, 0), 3))
        End If
    Next
End With
bExec = True
cmbBanco.ListIndex = 0

End Sub


Private Sub CallPb(nPos As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents
If nTotal = 0 Then Exit Sub
If ((nPos * 100) / nTotal) <= 100 Then
   PBar.value = (nPos * 100) / nTotal
Else
   PBar.value = 100
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

