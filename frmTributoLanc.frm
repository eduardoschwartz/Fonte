VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTributoLanc 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Tributos/Lancamentos"
   ClientHeight    =   4665
   ClientLeft      =   4020
   ClientTop       =   2625
   ClientWidth     =   6540
   Icon            =   "frmTributoLanc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstTrib 
      Height          =   3660
      Left            =   30
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   480
      Width           =   6465
   End
   Begin VB.ComboBox cmbLanc 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   5235
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5340
      TabIndex        =   3
      ToolTipText     =   "Sair da Tela"
      Top             =   4290
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
      MICON           =   "frmTributoLanc.frx":014A
      PICN            =   "frmTributoLanc.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   4260
      TabIndex        =   4
      ToolTipText     =   "Gravar os Dados"
      Top             =   4290
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmTributoLanc.frx":01D4
      PICN            =   "frmTributoLanc.frx":01F0
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
      Caption         =   "Lançamento..:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   150
      Width           =   1125
   End
End
Attribute VB_Name = "frmTributoLanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim sRet As String

Private Sub cmbLanc_Click()
CarregaTrib
End Sub

Private Sub cmdGravar_Click()
Dim x As Integer

Sql = "DELETE FROM TRIBUTOLANCAMENTO WHERE CODLANCAMENTO=" & cmbLanc.ItemData(cmbLanc.ListIndex)
cn.Execute Sql, rdExecDirect

For x = 0 To lstTrib.ListCount - 1
    If lstTrib.Selected(x) = True Then
        lstTrib.ListIndex = x
        Sql = "INSERT TRIBUTOLANCAMENTO(CODLANCAMENTO,CODTRIBUTO) VALUES("
        Sql = Sql & cmbLanc.ItemData(cmbLanc.ListIndex) & "," & lstTrib.ItemData(lstTrib.ListIndex) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next

MsgBox "Tributos gravados com sucesso.", vbInformation, "Informação"
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
Liberado
End Sub

Private Sub Form_Load()
Centraliza Me
sRet = RetEventUserForm(Me.Name)
CarregaLancamento

End Sub

Private Sub CarregaLancamento()

Sql = "Select CODLANCAMENTO,DESCFULL From LANCAMENTO ORDER BY DESCFULL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

cmbLanc.Clear
With RdoAux
   .MoveFirst
    Do Until .EOF
       cmbLanc.AddItem !DESCFULL
       cmbLanc.ItemData(cmbLanc.NewIndex) = !CodLancamento
      .MoveNext
    Loop
   .Close
End With
cmbLanc.ListIndex = 0

Sql = "Select CODTRIBUTO,DESCTRIBUTO From TRIBUTO ORDER BY DESCTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

lstTrib.Clear
With RdoAux
   .MoveFirst
    Do Until .EOF
       lstTrib.AddItem !DESCTRIBUTO
       lstTrib.ItemData(lstTrib.NewIndex) = !CodTributo
      .MoveNext
    Loop
   .Close
End With
cmbLanc_Click
End Sub

Private Sub CarregaTrib()
Dim x As Integer

Sql = "SELECT TRIBUTO.CODTRIBUTO,DESCTRIBUTO,ABREVTRIBUTO FROM TRIBUTO INNER JOIN "
Sql = Sql & "TRIBUTOLANCAMENTO ON TRIBUTO.CODTRIBUTO = TRIBUTOLANCAMENTO.CODTRIBUTO "
Sql = Sql & "WHERE TRIBUTOLANCAMENTO.CODLANCAMENTO =" & cmbLanc.ItemData(cmbLanc.ListIndex)
Sql = Sql & " ORDER BY DESCTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

For x = 0 To lstTrib.ListCount - 1
   lstTrib.Selected(x) = False
Next

With RdoAux
    Do Until .EOF
       For x = 0 To lstTrib.ListCount - 1
           If lstTrib.ItemData(x) = !CodTributo Then
              lstTrib.Selected(x) = True
              Exit For
           End If
       Next
      .MoveNext
    Loop
   .Close
End With

End Sub
