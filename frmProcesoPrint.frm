VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProcesoPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão"
   ClientHeight    =   3240
   ClientLeft      =   5520
   ClientTop       =   4275
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   300
      Left            =   1575
      TabIndex        =   2
      ToolTipText     =   "Imprimir relatório"
      Top             =   2835
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   529
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
      MICON           =   "frmProcesoPrint.frx":0000
      PICN            =   "frmProcesoPrint.frx":001C
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
      Height          =   285
      Left            =   270
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   2850
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   503
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
      MICON           =   "frmProcesoPrint.frx":0176
      PICN            =   "frmProcesoPrint.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstCol 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   45
      Width           =   2895
   End
End
Attribute VB_Name = "frmProcesoPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
Dim ax As String, z As Long, x As Integer, y As Integer, sChar As String, bAchou As Boolean

bAchou = False
For x = 0 To lstCol.ListCount - 1
    If lstCol.Selected(x) = True Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    MsgBox "Selecione ao menos uma coluna.", vbExclamation, "Atenção"
    Exit Sub
End If

sChar = ","

Ocupado
Open sPathBin & "\PROCESSOS.CSV" For Output As #1

ax = ""
With frmCnsProcesso2.grdProc
    For y = 1 To .Columns
        If lstCol.Selected(y - 1) = True Then
            ax = ax & .ColumnHeader(y) & sChar
        End If
    Next
    ax = Chomp(ax, chomp_righT, 1)
    Print #1, ax
   
    For x = 1 To .Rows
         ax = ""
        For y = 1 To .Columns
            If lstCol.Selected(y - 1) = True Then
                ax = ax & .CellText(x, y) & sChar
            End If
        Next
        ax = Chomp(ax, chomp_righT, 1)
        Print #1, ax
    Next
End With


Close #1
Liberado
MsgBox "O arquivo foi salvo em " & sPathBin & "\PROCESSOS.CSV"

Unload Me

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer

With frmCnsProcesso2.grdProc
    For x = 1 To .Columns
        lstCol.AddItem .ColumnHeader(x)
    Next
End With

End Sub
