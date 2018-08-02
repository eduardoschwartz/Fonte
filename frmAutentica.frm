VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAutentica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autenticação de Certidão"
   ClientHeight    =   2970
   ClientLeft      =   4725
   ClientTop       =   3555
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   7740
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2490
      Width           =   3435
   End
   Begin VB.TextBox txtChave 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   1020
      TabIndex        =   5
      Top             =   2100
      Width           =   6555
   End
   Begin VB.TextBox txtChave 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   1020
      TabIndex        =   4
      Top             =   1710
      Width           =   6555
   End
   Begin VB.TextBox txtChave 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   1020
      TabIndex        =   3
      Top             =   1320
      Width           =   6555
   End
   Begin VB.TextBox txtChave 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   1020
      TabIndex        =   2
      Top             =   930
      Width           =   6555
   End
   Begin VB.TextBox txtChave 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   540
      Width           =   6555
   End
   Begin VB.TextBox txtChave 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   150
      Width           =   6555
   End
   Begin prjChameleon.chameleonButton cmdValidar 
      Default         =   -1  'True
      Height          =   315
      Left            =   6210
      TabIndex        =   7
      Top             =   2520
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Validar"
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
      MICON           =   "frmAutentica.frx":0000
      PICN            =   "frmAutentica.frx":001C
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
      Caption         =   "Senha..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   90
      TabIndex        =   14
      Top             =   2550
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Chave 6:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   90
      TabIndex        =   13
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Chave 5:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   1770
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Chave 4:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   1380
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Chave 3:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Top             =   990
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Chave 2:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Chave 1:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   210
      Width           =   885
   End
End
Attribute VB_Name = "frmAutentica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdValidar_Click()
Dim sChave1 As String, sChave2 As String, sChave3 As String, sChave4 As String, sChave5 As String, sChave6 As String, sChave As String

sChave = Chr(75) & Chr(79) & Chr(66) & Chr(85) & Chr(68) & Chr(69) & Chr(82) & Chr(65)
For x = 0 To 5
    If Trim(txtChave(x)) = "" Then
        MsgBox "Favor preencher todas as chaves.", vbInformation, "Atenção"
        Exit Sub
    End If
Next

If UCase(txtSenha.text) <> Chr(80) & Chr(65) & Chr(67) & Chr(73) & Chr(70) & Chr(73) & Chr(67) & Chr(79) Then
    MsgBox "Senha inválida.", vbCritical, "Atenção"
    Exit Sub
End If

sChave1 = Decrypt128(txtChave(0).text, sChave)
sChave2 = Decrypt128(txtChave(1).text, sChave)
sChave3 = Decrypt128(txtChave(2).text, sChave)
sChave4 = Decrypt128(txtChave(3).text, sChave)
sChave5 = Decrypt128(txtChave(4).text, sChave)
sChave6 = Decrypt128(txtChave(5).text, sChave)


If Not IsDate(sChave3) Then
    MsgBox "Chave 3 inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If Right$(sChave4, 5) <> Mid(sChave3, 1, 5) Then
    MsgBox "Chave 4 inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(sChave6) = 0 Then
    MsgBox "Chave 6 inválida.", vbExclamation, "Atenção"
    Exit Sub
End If


txtChave(0).text = "Nome do usuário: " & sChave1
txtChave(1).text = "Nome da máquina: " & sChave2
txtChave(2).text = "Data de Emissão: " & sChave3
txtChave(3).text = "Numero Certidão: " & Left$(sChave4, Len(sChave4) - 5)
txtChave(4).text = "Tipo Certidão..: " & sChave5
txtChave(5).text = "Código Reduzido: " & sChave6

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub txtChave_KeyPress(Index As Integer, KeyAscii As Integer)
Tweak txtChave(Index), KeyAscii, AllLettersAllSmall
End Sub
