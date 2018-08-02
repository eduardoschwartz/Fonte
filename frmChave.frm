VERSION 5.00
Begin VB.Form frmChave 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chave de Acesso"
   ClientHeight    =   885
   ClientLeft      =   4770
   ClientTop       =   3660
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   285
      Left            =   5085
      TabIndex        =   3
      Top             =   450
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   90
      Picture         =   "frmChave.frx":0000
      ScaleHeight     =   510
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   135
      Width           =   465
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   285
      Left            =   5085
      TabIndex        =   1
      Top             =   135
      Width           =   915
   End
   Begin VB.TextBox txtPas 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1125
      TabIndex        =   0
      Top             =   180
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   495
      Picture         =   "frmChave.frx":030A
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "frmChave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

Dim RdoAux As rdoResultset, Sql As String

Dim volume_name As String
Dim serial_number As Long
Dim max_component_length As Long
Dim file_system_flags As Long
Dim file_system_name As String
Dim sCryptoKey As String

If GetVolumeInformation("C:\", volume_name, _
    Len(volume_name), serial_number, _
    max_component_length, file_system_flags, _
    file_system_name, Len(file_system_name)) = 0 _
Then
    MsgBox "No Disk In Drive!", vbInformation, "Error Reading Disk"
'    Security = False
    Exit Sub
End If

Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='CODDATAANT'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)

If KeyGen(Left$(NomeDoComputador, 2) & serial_number, RdoAux!VALPARAM, 3) <> txtPas.text Then
   MsgBox "Chave inserida inválida.", vbCritical, "Acesso negado ao Sistema !!!"
   Exit Sub
Else

   Sql = "DELETE FROM PARAMETROS WHERE NOMEPARAM='" & Left$(NomeDoComputador, 2) & serial_number & "'"
   cn.Execute Sql, rdExecDirect

   sCryptoKey = Encrypt128(txtPas.text, "sysTribut")
   Sql = "INSERT PARAMETROS(NOMEPARAM,VALPARAM) VALUES('"
   Sql = Sql & Left$(NomeDoComputador, 2) & serial_number & "','" & sCryptoKey & "')"
   cn.Execute Sql, rdExecDirect
End If

Unload Me
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Centraliza Me
RemoveTitleBar Me

End Sub


