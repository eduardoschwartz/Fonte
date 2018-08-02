VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmAgenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda"
   ClientHeight    =   5910
   ClientLeft      =   7095
   ClientTop       =   4065
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   13005
   Begin VB.ComboBox cmbExibir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmAgenda.frx":0000
      Left            =   210
      List            =   "frmAgenda.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5400
      Width           =   3285
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   555
      Left            =   6630
      TabIndex        =   4
      ToolTipText     =   "Novo Registro"
      Top             =   5250
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Alterar tarefa selecionada"
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
      MICON           =   "frmAgenda.frx":0051
      PICN            =   "frmAgenda.frx":006D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDesmarcar 
      Height          =   555
      Left            =   10590
      TabIndex        =   3
      ToolTipText     =   "Marcar como concluido"
      Top             =   5250
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Desmarcar como concluido"
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
      MICON           =   "frmAgenda.frx":010D
      PICN            =   "frmAgenda.frx":0129
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdMarcar 
      Height          =   555
      Left            =   8610
      TabIndex        =   2
      ToolTipText     =   "Marcar como concluido"
      Top             =   5250
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Marcar como concluido"
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
      MICON           =   "frmAgenda.frx":065B
      PICN            =   "frmAgenda.frx":0677
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
      Height          =   555
      Left            =   4650
      TabIndex        =   1
      ToolTipText     =   "Novo Registro"
      Top             =   5250
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   979
      BTYPE           =   3
      TX              =   "Criar uma nova tarefa"
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
      MICON           =   "frmAgenda.frx":0B99
      PICN            =   "frmAgenda.frx":0BB5
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
      Height          =   4965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   8758
      GridLines       =   -1  'True
      NoHorizontalGridLines=   -1  'True
      BackgroundPictureHeight=   300
      BackgroundPictureWidth=   300
      BackColor       =   16777215
      GridLineColor   =   8421504
      AlternateRowBackColor=   13697023
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
      HeaderHeight    =   25
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      DrawFocusRectangle=   0   'False
      DefaultRowHeight=   -1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exibir tarefas:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   1425
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNovo_Click()
frmTarefa.show vbModal
End Sub

Private Sub Form_Load()
Centraliza Me
cmbExibir.ListIndex = 0
GridHeader
LoadAgenda
End Sub

Private Sub GridHeader()
With grdMain
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    
    .AddColumn "kId", "Id", ecgHdrTextALignRight, , 20, False
    .AddColumn "kDti", "Dt.Incl.", ecgHdrTextALignCentre, , 70
    .AddColumn "kDtf", "Dt.Final", ecgHdrTextALignCentre, , 70
    .AddColumn "kUse", "Incluido por", ecgHdrTextALignLeft, , 180
    .AddColumn "kCom", "Descrição da tarefa", ecgHdrTextALignLeft, , 400
    .AddColumn "kfal", "Faltam", ecgHdrTextALignCentre, , 50
    .AddColumn "kDcn", "Dt.Concl.", ecgHdrTextALignCentre, , 70
End With

End Sub

Public Sub LoadAgenda()
Dim sql As String, RdoAux As rdoResultset, x As Integer

grdMain.Clear
sql = "select * from agenda where user_receive='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !id, DT_RIGHT
        grdMain.CellDetails grdMain.Rows, 2, Format(!data_inclusao, "dd/mm/yyyy"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 3, Format(!data_previsao, "dd/mm/yyyy"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 4, RetornaUsuarioFullName2(!user_send), DT_WORDBREAK
        grdMain.CellDetails grdMain.Rows, 5, !compromisso, DT_WORDBREAK
        grdMain.CellDetails grdMain.Rows, 6, "12", DT_CENTER
        grdMain.CellDetails grdMain.Rows, 7, "-----", DT_CENTER
       .MoveNext
    Loop
   .Close
End With

For x = 1 To grdMain.Rows
    grdMain.AutoHeightRow x
Next

grdMain.HighlightBackColor = &HFFFFC0
'grdMain.HighlightBackColor = &HD0FFFF
grdMain.HighlightForeColor = vbBlack



End Sub

