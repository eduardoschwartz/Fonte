VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCnsNumProc 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Nº de Processo"
   ClientHeight    =   3840
   ClientLeft      =   2115
   ClientTop       =   2055
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   6435
   Begin VB.ComboBox cmbAno 
      Height          =   315
      Left            =   2460
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   210
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid grdProc 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   12582912
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Nº Processo        |^Data Processo     |^Cancelado  |^Data Cancelado   |^Cód.            "
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "Cancelar Edição"
      Top             =   3300
      Width           =   1305
      _ExtentX        =   2302
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
      MICON           =   "frmCnsNumProc.frx":0000
      PICN            =   "frmCnsNumProc.frx":001C
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
      Left            =   3630
      TabIndex        =   4
      ToolTipText     =   "Cadastra o Imóvel"
      Top             =   3300
      Width           =   1305
      _ExtentX        =   2302
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
      MICON           =   "frmCnsNumProc.frx":0176
      PICN            =   "frmCnsNumProc.frx":0192
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
      Caption         =   "Selecione o Ano do Processo:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   2235
   End
End
Attribute VB_Name = "frmCnsNumProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Dim RdoAux As rdoResultset

Private Sub cmbAno_Click()
CarregaLista
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRetorna_Click()
Dim nCodImovel As Long
Dim sNumProc As String

If grdProc.Rows > 1 Then
    If grdProc.Row > 0 Then
       nCodImovel = grdProc.TextMatrix(grdProc.Row, 4)
       frmCancelReparc.txtCod.text = nCodImovel
       frmCancelReparc.txtCod_LostFocus
       frmCancelReparc.cmbProc.SetFocus
       Sql = "SELECT NUMPROCESSO FROM VWNUMPROCESSO WHERE ANOPROC=" & cmbAno.text & " AND CONVERT(INT,NUMERO)=" & Val(grdProc.TextMatrix(grdProc.Row, 0))
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
            sNumProc = !NUMPROCESSO
           .Close
       End With
       If frmCancelReparc.cmbProc.ListCount > 0 Then
          frmCancelReparc.cmbProc.text = sNumProc
       End If
    Else
       Exit Sub
    End If
Else
    Exit Sub
End If
Unload Me
End Sub

Private Sub Form_Load()

Ocupado

Liberado

Sql = "SELECT DISTINCT ANOPROC FROM VWNUMPROCESSO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAno.AddItem !anoproc
       .MoveNext
    Loop
    If cmbAno.ListCount > 0 Then cmbAno.ListIndex = 0
End With
End Sub

Private Sub CarregaLista()

grdProc.Rows = 1
Sql = "SELECT NUMERO,DATAPROCESSO,CANCELADO,DATACANCEL,CODIGORESP FROM VWNUMPROCESSO WHERE ANOPROC=" & cmbAno.text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdProc.AddItem Format(!Numero, "000000") & Chr(9) & Format(!DATAPROCESSO, "dd/mm/yyyy") & Chr(9) & IIf(!Cancelado, "Sim", "Não") & Chr(9) & !DATACANCEL & Chr(9) & Format(!CODIGORESP, "00000")
       .MoveNext
    Loop
End With

End Sub
