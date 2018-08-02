VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSIL 
   BackColor       =   &H00808000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SIL"
   ClientHeight    =   2640
   ClientLeft      =   5385
   ClientTop       =   3900
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSil 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6825
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      ToolTipText     =   "Excluir Registro"
      Top             =   2220
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmSIL.frx":0000
      PICN            =   "frmSIL.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   2220
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmSIL.frx":00BE
      PICN            =   "frmSIL.frx":00DA
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
      Height          =   315
      Left            =   180
      TabIndex        =   3
      ToolTipText     =   "Novo Registro"
      Top             =   2220
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmSIL.frx":0234
      PICN            =   "frmSIL.frx":0250
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
      Left            =   5700
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   2220
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
      MICON           =   "frmSIL.frx":03AA
      PICN            =   "frmSIL.frx":03C6
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
Attribute VB_Name = "frmSIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodReduz As Long

Private Sub cmdAlterar_Click()
If lstSil.ListIndex = -1 Then
    MsgBox "Nada a alterar.", vbCritical, "Erro"
    Exit Sub
End If

Dim z As Variant
z = InputBox("Digite o novo SIL", "Alteração de SIL", lstSil.Text)
If SubNull(z) <> "" Then
    Sql = "update sil set sil='" & Mask(CStr(z)) & "' where sid=" & lstSil.ItemData(lstSil.ListIndex)
    cn.Execute Sql, rdExecDirect
    LoadLista
End If

End Sub

Private Sub cmdExcluir_Click()
If lstSil.ListIndex = -1 Then
    MsgBox "Nada a excluir.", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Excluir este Sil?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "delete from sil where sid=" & lstSil.ItemData(lstSil.ListIndex)
    cn.Execute Sql, rdExecDirect
    LoadLista
End If

End Sub

Private Sub cmdNovo_Click()
Dim z As Variant
z = InputBox("Digite o novo SIL", "Inclusão de SIL")
If SubNull(z) <> "" Then
    Sql = "insert sil(codigo,sil) values(" & nCodReduz & ",'" & Mask(CStr(z)) & "')"
    cn.Execute Sql, rdExecDirect
    LoadLista
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
nCodReduz = Val(frmCadMob.txtCodEmpresa.Text)
LoadLista
End Sub

Private Sub LoadLista()
Dim Sql As String, RdoAux As rdoResultset
lstSil.Clear
Sql = "select * from sil where codigo=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstSil.AddItem !sil
        lstSil.ItemData(lstSil.NewIndex) = !sID
       .MoveNext
    Loop
   .Close
End With
If lstSil.ListCount > 0 Then lstSil.ListIndex = 0

End Sub

