VERSION 5.00
Begin VB.Form frmMonitor 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monitoração"
   ClientHeight    =   1035
   ClientLeft      =   3000
   ClientTop       =   4140
   ClientWidth     =   4200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   120
      Width           =   2865
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   540
      Width           =   2865
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   570
      Width           =   1005
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnMon As New rdoConnection, RdoAux As rdoResultset, Sql As String

Private Sub Form_KeyPress(KeyAscii As Integer)


If KeyAscii = vbKeyReturn Then
    txtUser.text = UCase$(txtUser.text)
    txtPwd.text = UCase$(txtPwd.text)
    If Monitor(txtUser.text, txtPwd.text) Then
        Sql = "SELECT * FROM SEG_USERACESS WHERE NOMEUSUARIO='" & txtUser.text & "' AND CODTELA=72 AND CODEVENTO=4"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Este usuário não é um Supervisor", vbCritical, "Atenção"
                Exit Sub
            End If
           .Close
        End With
        frmCancelDebito.lblsupervisor.Caption = txtUser.text
        frmCancelDebito.lblSup.Caption = "1"
        cnMon.Close
        Unload Me
    Else
        MsgBox "Usuário/Senha Inválido.", vbCritical, "Atenção"
        frmCancelDebito.lblSup.Caption = "0"
    End If
ElseIf KeyAscii = vbKeyEscape Then
    frmCancelDebito.lblSup.Caption = "0"
    Unload Me
End If

End Sub

Private Function Monitor(User As String, Pwd As String) As Boolean

On Error GoTo Erro
Dim LoginDSN As String

LoginDSN = "odbcTributacao"

If Trim$(User) = "" Then
     Monitor = False
     Exit Function
End If

Set en = rdoEngine.rdoEnvironments(0)
en.CursorDriver = rdUseServer
With en
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 20
     
    Set cnMon = en.OpenConnection(dsname:=LoginDSN, _
        Prompt:=rdDriverNoPrompt, _
        Connect:="uid=" & User & ";PWD=" & Pwd & ";driver={SQL Server};")
     
End With
Monitor = True

Exit Function
Erro:
Monitor = False

End Function


