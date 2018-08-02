VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmFiltroDebito 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtro de Débitos"
   ClientHeight    =   4905
   ClientLeft      =   4680
   ClientTop       =   2445
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbE 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   900
      Width           =   1065
   End
   Begin VB.ComboBox cmbL 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Width           =   2565
   End
   Begin VB.ComboBox cmbS 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2220
      Width           =   2595
   End
   Begin VB.ComboBox cmbD 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3600
      Width           =   1545
   End
   Begin VB.ComboBox cmbA 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4260
      Width           =   1545
   End
   Begin VB.ComboBox cmbSeq 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2910
      Width           =   1545
   End
   Begin VB.ComboBox cmbE2 
      Height          =   315
      Left            =   1620
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   1065
   End
   Begin prjChameleon.chameleonButton cmdFilter 
      Height          =   375
      Left            =   1830
      TabIndex        =   8
      ToolTipText     =   "Aplicar Filtro"
      Top             =   4080
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFiltroDebito.frx":0000
      PICN            =   "frmFiltroDebito.frx":001C
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
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      ToolTipText     =   "Fechar a Tela"
      Top             =   4080
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFiltroDebito.frx":01F6
      PICN            =   "frmFiltroDebito.frx":0212
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmFiltroDebito.frx":036C
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exercício:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   630
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lançamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   15
      Top             =   1290
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Situação Parcela:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   14
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Divida Ativa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   150
      TabIndex        =   13
      Top             =   3330
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ajuizado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   12
      Top             =   3990
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sequência Parcelamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   150
      TabIndex        =   11
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "à"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   1260
      TabIndex        =   10
      Top             =   960
      Width           =   285
   End
End
Attribute VB_Name = "frmFiltroDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExec As Boolean, Sql As String, RdoAux As rdoResultset


Private Sub cmdFilter_Click()
If Not bExec Then Exit Sub
If cmbA.ListIndex = 0 Then
   FiltroA = "X"
Else
   FiltroA = Left$(cmbA.text, 1)
End If

If cmbE.ListIndex = 0 Then
   FiltroE = 0
Else
   FiltroE = cmbE.text
End If
If cmbE2.ListIndex = 0 Or cmbE2.ListIndex = -1 Then
   FiltroE2 = 0
Else
   FiltroE2 = cmbE2.text
End If

If cmbL.ListIndex = 0 Then
   FiltroL = 0
   FiltroLP = ""
Else
   FiltroL = Left$(cmbL.text, 3)
   If InStr(1, cmbL.text, "(") > 0 Then
      FiltroLP = Mid(cmbL.text, InStr(1, cmbL.text, "(") + 1, Len(cmbL.text) - InStr(1, cmbL.text, "(") - 1)
      If InStr(1, FiltroLP, "-", vbBinaryCompare) > 0 Then
        FiltroLP = Left$(FiltroLP, InStr(1, FiltroLP, "/", vbBinaryCompare) - 3) & "/" & Right$(FiltroLP, 4)
      End If
   Else
      FiltroLP = ""
   End If
End If

If cmbSeq.ListIndex = 0 Then
   FiltroSEQ = 0
   bFiltroSEQ = False
Else
   bFiltroSEQ = True
   FiltroSEQ = Left$(cmbSeq.text, 3)
End If

If cmbS.ListIndex = 0 Then
   FiltroS = 99
Else
   FiltroS = Left$(cmbS.text, 2)
End If
If cmbD.ListIndex = 0 Then
   FiltroD = "X"
Else
   FiltroD = Left$(cmbD.text, 1)
End If

frmDebitoImob.CarregaDebito2 (Val(frmDebitoImob.txtCod.text))

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim v As Integer, w As Integer, Achou As Boolean

Me.Left = frmMdi.Width - Me.Width - 200
bExec = False
cmbA.AddItem "(Todos)"
cmbA.AddItem "Sim"
cmbA.AddItem "Não"
Select Case FiltroA
    Case "X"
      cmbA.ListIndex = 0
    Case "S"
      cmbA.ListIndex = 1
    Case "N"
      cmbA.ListIndex = 2
End Select

cmbD.AddItem "(Todos)"
cmbD.AddItem "Sim"
cmbD.AddItem "Não"
Select Case FiltroD
    Case "X"
      cmbD.ListIndex = 0
    Case "S"
      cmbD.ListIndex = 1
    Case "N"
      cmbD.ListIndex = 2
End Select

cmbL.AddItem "(Todos)"
For v = 1 To frmDebitoImob.grdExtrato.Rows
   Achou = False
   For w = 0 To cmbL.ListCount - 1
       cmbL.ListIndex = w
       If frmDebitoImob.grdExtrato.CellText(v, 2) = cmbL.text Then
          Achou = True
       End If
   Next
   If Not Achou Then
      cmbL.AddItem frmDebitoImob.grdExtrato.CellText(v, 2)
   End If
Next
Select Case FiltroL
    Case 0
      cmbL.ListIndex = 0
    Case Else
      For v = 0 To cmbL.ListCount - 1
          cmbL.ListIndex = v
          If Val(Left$(cmbL.text, 3)) = FiltroL Then
             Exit For
          End If
      Next
End Select


cmbS.AddItem "(Todos)"
For v = 1 To frmDebitoImob.grdExtrato.Rows
   Achou = False
   For w = 0 To cmbS.ListCount - 1
       cmbS.ListIndex = w
       If frmDebitoImob.grdExtrato.CellText(v, 6) = cmbS.text Then
          Achou = True
       End If
   Next
   If Not Achou Then
      cmbS.AddItem frmDebitoImob.grdExtrato.CellText(v, 6)
   End If
Next
Select Case FiltroS
    Case 99
      cmbS.ListIndex = 0
    Case Else
      For v = 0 To cmbS.ListCount - 1
          cmbS.ListIndex = v
          If Val(Left$(cmbS.text, 2)) = FiltroS Then
             Exit For
          End If
      Next
End Select

cmbSeq.AddItem "(Todos)"
Sql = "SELECT DISTINCT SEQLANCAMENTO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(frmDebitoImob.txtCod.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbSeq.AddItem !SeqLancamento
       .MoveNext
    Loop
   .Close
End With
cmbSeq.ListIndex = 0

cmbE.AddItem "(Todos)"
cmbE2.AddItem "(Todos)"
For v = 1 To frmDebitoImob.grdExtrato.Rows
   Achou = False
   For w = 0 To cmbE.ListCount - 1
       cmbE.ListIndex = w
       If frmDebitoImob.grdExtrato.CellText(v, 1) = cmbE.text Then
          Achou = True
       End If
   Next
   If Not Achou Then
      cmbE.AddItem frmDebitoImob.grdExtrato.CellText(v, 1)
      cmbE2.AddItem frmDebitoImob.grdExtrato.CellText(v, 1)
   End If
Next
Select Case FiltroE
    Case 0
      cmbE.ListIndex = 0
    Case Else
      For v = 0 To cmbE.ListCount - 1
          cmbE.ListIndex = v
          If Val(Left$(cmbE.text, 2)) = FiltroE Then
             Exit For
          End If
      Next
End Select

bExec = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

frmDebitoImob.cmdFilter.Value = False
End Sub

