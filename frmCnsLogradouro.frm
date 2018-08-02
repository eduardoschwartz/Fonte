VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCnsLogradouro 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Logradouro"
   ClientHeight    =   4920
   ClientLeft      =   3975
   ClientTop       =   2235
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      ToolTipText     =   "Selecionar Logradouro"
      Top             =   4530
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Selecionar"
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
      MICON           =   "frmCnsLogradouro.frx":0000
      PICN            =   "frmCnsLogradouro.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Nome"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   3390
      TabIndex        =   4
      Top             =   4560
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Tít&ulo"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   2310
      TabIndex        =   3
      Top             =   4560
      Width           =   1005
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Tipo"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1230
      TabIndex        =   2
      Top             =   4560
      Width           =   1005
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Código"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   4560
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid grdLogr 
      Height          =   4380
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   7726
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   ">Código   |<Tipo             |<Título                |<Nome do Logradouro                                                       "
   End
End
Attribute VB_Name = "frmCnsLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String


Private Sub cmdSair_Click()

For x = 0 To Forms.Count - 1
    If Forms(x).Name = "frmCnsImovel" Then
       frmCnsImovel.txtCodLogr.Text = Format(grdLogr.TextMatrix(grdLogr.Row, 0), "000000")
       Unload frmCnsLogradouro
       frmCnsImovel.SetFocus
       Exit Sub
    End If
Next
Unload Me

End Sub

Private Sub Form_Load()
Ocupado

CarregaLista 3
Centraliza Me
Liberado
End Sub

Private Sub CarregaLista(nOrdem As Integer)

Screen.MousePointer = vbHourglass

Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
Select Case nOrdem
   Case 0
      Sql = Sql & "ORDER BY CODLOGRADOURO"
   Case 1
      Sql = Sql & "ORDER BY ABREVTIPOLOG"
   Case 2
      Sql = Sql & "ORDER BY ABREVTITLOG"
   Case 3
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
End Select

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
grdLogr.Rows = 1
With RdoAux
   .MoveFirst
    Do Until .EOF
       grdLogr.AddItem !CodLogradouro & Chr(9) & !AbrevTipoLog & Chr(9) & SubNull(!AbrevTitLog) & Chr(9) & !NomeLogradouro
      .MoveNext
    Loop
   .Close
End With

grdLogr.col = 0
grdLogr.ColSel = 3

Screen.MousePointer = vbDefault

End Sub

Private Sub Opt_Click(Index As Integer)
Ocupado
CarregaLista Index
Liberado
End Sub


