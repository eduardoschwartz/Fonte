VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmNF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Fiscais Emitidas"
   ClientHeight    =   4305
   ClientLeft      =   3060
   ClientTop       =   2955
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4080
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   2910
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   3900
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmNF.frx":0000
      PICN            =   "frmNF.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvNF 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      _ExtentX        =   7170
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código Emp"
         Object.Width           =   2188
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Nota Fiscal"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Data Aut."
         Object.Width           =   2187
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdBaixa 
      Height          =   345
      Left            =   1755
      TabIndex        =   2
      Top             =   3900
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmNF.frx":008A
      PICN            =   "frmNF.frx":00A6
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
Attribute VB_Name = "frmNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset

Private Sub cmdBaixa_Click()
frmReport.ShowReport "MOBILIARIONF", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Carrega
End Sub

Private Sub Carrega()
Dim itmX As ListItem, z As Long
Ocupado
z = SendMessage(lvNF.hwnd, LVM_DELETEALLITEMS, 0, 0)
On Error Resume Next
Sql = "SELECT CODIGOMOB, NUMAUT,DATAAUT From MOBILIARIONF WHERE NUMAUT <> 'XXXX' ORDER BY NUMAUT"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Set itmX = lvNF.ListItems.Add(, "NF" & Format(!CODIGOMOB, "000000") & Format(!NUMAUT, "0000"), !CODIGOMOB)
       itmX.SubItems(1) = Format(Trim$(!NUMAUT), "0000")
       If Not IsNull(!DATAAUT) Then
        itmX.SubItems(2) = Format(!DATAAUT, "dd/mm/yyyy")
       End If
      .MoveNext
    Loop
   .Close
End With
Liberado

End Sub

Private Sub lvNF_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvNF.SortKey = ColumnHeader.Position - 1
lvNF.Sorted = True
lvNF.SortOrder = lvwAscending

End Sub
