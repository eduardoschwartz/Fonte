VERSION 5.00
Begin VB.Form frmHist 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico do Sistema"
   ClientHeight    =   3210
   ClientLeft      =   6585
   ClientTop       =   4305
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   7485
   Begin VB.TextBox txtHist 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   510
      Left            =   45
      TabIndex        =   1
      Top             =   2655
      Width           =   7395
   End
   Begin VB.ListBox lstLog 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7395
   End
End
Attribute VB_Name = "frmHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Centraliza Me
Ocupado
Le
lstLog_Click
Liberado
End Sub

Private Sub Le()
Dim FF1 As Integer, sReg As String, aRegistro() As String

sPathBin = App.Path & "\bin"

If Dir(sPathBin & "\gti.000") = "" Then
    FF1 = FreeFile()
    Open sPathBin & "\gti.000" For Output As FF1
    Close #FF1
    Exit Sub
End If

FF1 = FreeFile()
Open sPathBin & "\gti.000" For Input As FF1
While Not EOF(FF1)
    Input #1, sReg
    sReg = Decrypt128(sReg, MBI_LG)
    aRegistro = Split(sReg, "#")
    lstLog.AddItem aRegistro(3) & " - " & aRegistro(4)
Wend
Close #FF1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Me.Hide
End Sub

Private Sub lstLog_Click()
txtHist.text = lstLog.text
End Sub
