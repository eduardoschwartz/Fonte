VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLog 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log do Sistema"
   ClientHeight    =   5895
   ClientLeft      =   1140
   ClientTop       =   2460
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   8520
      TabIndex        =   21
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   570
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmLog.frx":0000
      PICN            =   "frmLog.frx":001C
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
      Left            =   8520
      TabIndex        =   22
      ToolTipText     =   "Sair da Tela"
      Top             =   990
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLog.frx":0176
      PICN            =   "frmLog.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdFilter 
      Height          =   315
      Left            =   8520
      TabIndex        =   23
      ToolTipText     =   "Sair da Tela"
      Top             =   150
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "F&iltrar"
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
      MICON           =   "frmLog.frx":0200
      PICN            =   "frmLog.frx":021C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00EEEEEE&
      Height          =   675
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   5190
      Width           =   9645
   End
   Begin MSComctlLib.ImageList ilsIcon 
      Left            =   5310
      Top             =   2790
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0376
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":04D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0636
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0796
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":08F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":1092
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvLog 
      Height          =   3135
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   5530
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilsIcon"
      SmallIcons      =   "ilsIcon"
      ColHdrIcons     =   "ilsIcon"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "H1"
         Text            =   "Computador"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "H2"
         Text            =   "Usuário"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "H3"
         Text            =   "Data e Hora"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "H4"
         Text            =   "Evento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "H5"
         Text            =   "SubEvento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "H6"
         Text            =   "Tela"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEEE&
      Height          =   1365
      Left            =   0
      TabIndex        =   9
      Top             =   660
      Width           =   8445
      Begin VB.ComboBox cmbComputer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   2835
      End
      Begin VB.ComboBox cmbUser 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   570
         Width           =   2835
      End
      Begin VB.ComboBox cmbEvento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   930
         Width           =   2835
      End
      Begin VB.ComboBox cmbSubEvento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2835
      End
      Begin VB.ComboBox cmbTela 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5490
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   2835
      End
      Begin VB.Label lblTotLog 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5520
         TabIndex        =   17
         Top             =   1020
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Logs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4440
         TabIndex        =   16
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Computador:"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Evento:"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Evento:"
         Height          =   225
         Index           =   5
         Left            =   4230
         TabIndex        =   11
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da Tela:"
         Height          =   225
         Index           =   6
         Left            =   4230
         TabIndex        =   10
         Top             =   660
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Período"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8445
      Begin MSComCtl2.DTPicker vcDataDe 
         Height          =   315
         Left            =   2970
         TabIndex        =   19
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   50593793
         CurrentDate     =   37594
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Todos os Períodos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   0
         Top             =   300
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker vcDataAte 
         Height          =   315
         Left            =   5550
         TabIndex        =   20
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   50593793
         CurrentDate     =   37594
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Período de:"
         Height          =   225
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Período até:"
         Height          =   225
         Index           =   1
         Left            =   4590
         TabIndex        =   7
         Top             =   300
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String

Private Sub chkTodos_Click()

If chkTodos.Value = 1 Then
   vcDataDe.Enabled = False
   vcDataAte.Enabled = False
Else
   vcDataDe.Enabled = True
   vcDataAte.Enabled = True
End If

End Sub

Private Sub cmdAbaixo_Click()

If cmdAbaixo.Value = Down Then
   cmdAcima.Value = Up
Else
   cmdAcima.Value = Down
End If

End Sub

Private Sub cmdAcima_Click()

If cmdAcima.Value = Down Then
   cmdAbaixo.Value = Up
Else
   cmdAbaixo.Value = Down
End If

End Sub

Private Sub cmdFilter_Click()
Dim itmX As ListItem
Dim bComputador As Boolean, bUsuario As Boolean, bEvento As Boolean
Dim bSecEvento As Boolean, bForm As Boolean, bData As Boolean
Dim lStyle As Long
Dim lR As Long, tLV As LVITEM
Dim z As Long

Ocupado
Screen.MousePointer = vbHourglass
lblTotLog.Caption = 0
txtLog.text = ""
z = SendMessage(lvLog.hwnd, LVM_DELETEALLITEMS, 0, 0)

If chkTodos.Value = 0 Then bData = True
If cmbComputer.ListIndex > 0 Then bComputador = True
If cmbUser.ListIndex > 0 Then bUsuario = True
If cmbEvento.ItemData(cmbEvento.ListIndex) <> 999 Then bEvento = True
If cmbSubEvento.ItemData(cmbSubEvento.ListIndex) <> 999 Then bSecEvento = True
If cmbTela.ListIndex > 0 Then bForm = True
   
Sql = "SELECT SEQ,DATAHORAEVENTO,COMPUTADOR,USUARIO,FORM,EVENTO,SECEVENTO FROM LOGEVENTO WHERE "

If bComputador Then
   Sql = Sql & "COMPUTADOR='" & cmbComputer.text & "' AND "
End If

If bUsuario Then
   Sql = Sql & "USUARIO='" & cmbUser.text & "' AND "
End If

If bForm Then
   Sql = Sql & "FORM='" & cmbTela.text & "' AND "
End If

If bEvento Then
   Sql = Sql & "EVENTO=" & cmbEvento.ItemData(cmbEvento.ListIndex) & " AND "
End If

If bSecEvento Then
   Sql = Sql & "SECEVENTO=" & cmbSubEvento.ItemData(cmbSubEvento.ListIndex) & " AND "
End If

If bData Then
   Sql = Sql & "CONVERT(CHAR(10),DATAHORAEVENTO,110) BETWEEN '" & Format(vcDataDe.Value, "mm-dd-yyyy") & "' AND '" & Format(vcDataAte.Value, "mm-dd-yyyy") & "'"
End If

If Right$(Sql, 6) = "WHERE " Then
   Sql = Left$(Sql, Len(Sql) - 6)
End If

If Right$(Sql, 4) = "AND " Then
   Sql = Left$(Sql, Len(Sql) - 4)
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
lblTotLog.Caption = RdoAux.RowCount
lStyle = SendMessageByLong(lvLog.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
SendMessageByLong lvLog.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle
tLV.Mask = LVIF_IMAGE

With lvLog
    Do While Not RdoAux.EOF
      'Texto
       Set itmX = .ListItems.Add(, "C" & Trim$(sTr(RdoAux!Seq)), RdoAux!COMPUTADOR)
      'Col2
       itmX.SubItems(1) = RdoAux!USUARIO
      'Col3
       itmX.SubItems(2) = RdoAux!DATAHORAEVENTO
      'Col4
       Select Case RdoAux!Evento
           Case 1
              itmX.SubItems(3) = "Logon"
           Case 2
              itmX.SubItems(3) = "Logoff"
           Case 3
              itmX.SubItems(3) = "Form"
           Case 4
              itmX.SubItems(3) = "Configuração"
           Case 999
              itmX.SubItems(3) = "(Todos)"
       End Select
       tLV.iSubItem = 3
       tLV.iItem = lvLog.ListItems.Count - 1
       Select Case RdoAux!Evento
           Case 1
              tLV.iImage = 5
           Case 2
              tLV.iImage = 6
           Case 3
              tLV.iImage = 7
       End Select
       SendMessage lvLog.hwnd, LVM_SETITEM, lR - 1, tLV
      
      'Col5
       Select Case RdoAux!SECEVENTO
           Case 0
              itmX.SubItems(4) = "Nenhum"
           Case 1
              itmX.SubItems(4) = "Inclusão"
           Case 2
              itmX.SubItems(4) = "Alteração"
           Case 3
              itmX.SubItems(4) = "Exclusão"
           Case 4
              itmX.SubItems(4) = "Impressão"
           Case 999
              itmX.SubItems(4) = "(Todos)"
       End Select

        tLV.iSubItem = 4
        tLV.iItem = lvLog.ListItems.Count - 1
        Select Case RdoAux!SECEVENTO
            Case 1
               tLV.iImage = 0
            Case 2
               tLV.iImage = 1
            Case 3
               tLV.iImage = 2
            Case Else
               tLV.iImage = 4
        End Select
        SendMessage lvLog.hwnd, LVM_SETITEM, lR - 1, tLV
    
     'Col6
       itmX.SubItems(5) = RdoAux!Form
       RdoAux.MoveNext
    Loop
End With
RdoAux.Close
Screen.MousePointer = vbDefault
Liberado
End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
bResize = True
End Sub

Private Sub Form_Load()

Ocupado
Centraliza Me
chkTodos.Value = 0
vcDataDe.Enabled = True
vcDataAte.Enabled = True
vcDataDe.Value = Format(Now, "dd/mm/yyyy")
vcDataAte.Value = Format(Now, "dd/mm/yyyy")
CarregaLista
Liberado

End Sub

Private Sub CarregaLista()

Sql = "SELECT DISTINCT COMPUTADOR FROM LOGEVENTO; " & _
      "SELECT DISTINCT USUARIO FROM LOGEVENTO WHERE USUARIO<>''; " & _
      "SELECT DISTINCT FORM FROM LOGEVENTO"
         
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    cmbComputer.AddItem "(Todos)"
    Do Until .EOF
       cmbComputer.AddItem !COMPUTADOR
      .MoveNext
    Loop
    cmbComputer.ListIndex = 0
   .MoreResults
    cmbUser.AddItem "(Todos)"
    Do Until .EOF
       cmbUser.AddItem !USUARIO
      .MoveNext
    Loop
    cmbUser.ListIndex = 0
   .MoreResults
    cmbTela.AddItem "(Todos)"
    Do Until .EOF
       cmbTela.AddItem !Form
      .MoveNext
    Loop
    cmbTela.ListIndex = 0
   .Close
End With

cmbEvento.AddItem "(Todos)"
cmbEvento.ItemData(cmbEvento.NewIndex) = 999
cmbEvento.AddItem "Logon"
cmbEvento.ItemData(cmbEvento.NewIndex) = 1
cmbEvento.AddItem "Logoff"
cmbEvento.ItemData(cmbEvento.NewIndex) = 2
cmbEvento.AddItem "Form"
cmbEvento.ItemData(cmbEvento.NewIndex) = 3
cmbEvento.AddItem "Configuração"
cmbEvento.ItemData(cmbEvento.NewIndex) = 4
cmbEvento.ListIndex = 0

cmbSubEvento.AddItem "(Todos)"
cmbSubEvento.ItemData(cmbSubEvento.NewIndex) = 999
cmbSubEvento.AddItem "Nenhum"
cmbSubEvento.ItemData(cmbSubEvento.NewIndex) = 0
cmbSubEvento.AddItem "Inclusão"
cmbSubEvento.ItemData(cmbSubEvento.NewIndex) = 1
cmbSubEvento.AddItem "Alteração"
cmbSubEvento.ItemData(cmbSubEvento.NewIndex) = 2
cmbSubEvento.AddItem "Exclusão"
cmbSubEvento.ItemData(cmbSubEvento.NewIndex) = 3
cmbSubEvento.AddItem "Impressão"
cmbSubEvento.ItemData(cmbSubEvento.NewIndex) = 4
cmbSubEvento.ListIndex = 0

End Sub

Private Sub lvLog_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim i As Long
Dim iT As Long

   ' Sort the ListView.  Demonstrates customised sorting routines

   ' Set tag saying which way we've sorted..
   If (ColumnHeader.Tag <> "ASC") Then
      m_eSortOrder = lvwAscending
      ColumnHeader.Tag = "ASC"
   Else
      m_eSortOrder = lvwDescending
      ColumnHeader.Tag = "DESC"
   End If

   ' Reset other items:
   iT = ColumnHeader.Index
   For i = 1 To 4
      If i <> iT Then
         lvLog.ColumnHeaders(i).Tag = ""
      End If
   Next i

   Select Case ColumnHeader.Key
   Case "H3"
      ' Date column - use custom sort
      lvLog.Sorted = False
      m_iSortCol = 3
      m_eSortType = elvstDate
      SendMessageByLong lvLog.hwnd, LVM_SORTITEMS, lvLog.hwnd, AddressOf LVWSortCompare
   Case Else
      ' Number column
      m_iSortCol = ColumnHeader.Index
      m_eSortType = elvstText
      SendMessageByLong lvLog.hwnd, LVM_SORTITEMS, lvLog.hwnd, AddressOf LVWSortCompare
   End Select
   
End Sub

Private Sub lvLog_ItemClick(ByVal Item As MSComctlLib.ListItem)

Sql = "SELECT LOGEVENTO From LOGEVENTO WHERE SEQ=" & Val(Right$(Item.Key, Len(Item.Key) - 1))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtLog.text = SubNull(!LOGEVENTO)
    Else
        txtLog.text = ""
    End If
   .Close
End With

End Sub

