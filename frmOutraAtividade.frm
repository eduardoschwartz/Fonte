VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmOutraAtividade 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outras Atividades"
   ClientHeight    =   2520
   ClientLeft      =   4995
   ClientTop       =   4245
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdHelp 
      Height          =   315
      Left            =   5310
      TabIndex        =   1
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   2100
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Ajuda"
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
      MICON           =   "frmOutraAtividade.frx":0000
      PICN            =   "frmOutraAtividade.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAdd 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Adicionar Atividade"
      Top             =   2100
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Adicionar"
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
      MICON           =   "frmOutraAtividade.frx":0176
      PICN            =   "frmOutraAtividade.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDel 
      Height          =   315
      Left            =   1305
      TabIndex        =   3
      ToolTipText     =   "Remover Atividade"
      Top             =   2100
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Remover"
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
      MICON           =   "frmOutraAtividade.frx":02EC
      PICN            =   "frmOutraAtividade.frx":0308
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
      Left            =   6525
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   2100
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmOutraAtividade.frx":0462
      PICN            =   "frmOutraAtividade.frx":047E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdAtiv 
      Height          =   1980
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   3493
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      Appearance      =   0
      FormatString    =   $"frmOutraAtividade.frx":04EC
   End
End
Attribute VB_Name = "frmOutraAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub cmdAdd_Click()
Set frm = frmAtiv
frm.sForm = "frmOutraAtividade"
frmAtiv.show vbModeless
frmAtiv.ZOrder 0
End Sub

Private Sub cmdDel_Click()
If grdAtiv.Rows = 1 Then
   MsgBox "Selecione a Atividade a ser excluída.", vbExclamation, "Atenção"
Else
   If grdAtiv.Rows > 2 Then
      grdAtiv.RemoveItem (grdAtiv.Row)
   Else
      grdAtiv.Rows = 1
   End If
End If

End Sub

Private Sub cmdSair_Click()
Dim x As Integer

frmCadMob.grdTemp.Rows = 1
For x = 1 To grdAtiv.Rows - 1
    With grdAtiv
        frmCadMob.grdTemp.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3) & Chr(9) & .TextMatrix(x, 4) & Chr(9) & .TextMatrix(x, 5)
    End With
Next

Unload frmAtiv
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer
Centraliza Me

grdAtiv.Rows = 1
For x = 1 To frmCadMob.grdTemp.Rows - 1
    With frmCadMob.grdTemp
        grdAtiv.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3) & Chr(9) & .TextMatrix(x, 4) & Chr(9) & .TextMatrix(x, 5)
    End With
Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CodEmpresa = 0
End Sub

