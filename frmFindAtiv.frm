VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmFindAtiv 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar"
   ClientHeight    =   1650
   ClientLeft      =   3075
   ClientTop       =   3165
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5325
   Begin VB.ComboBox cmbFind 
      Height          =   315
      ItemData        =   "frmFindAtiv.frx":0000
      Left            =   120
      List            =   "frmFindAtiv.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1170
      Width           =   3855
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   3855
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4110
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   1200
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmFindAtiv.frx":0061
      PICN            =   "frmFindAtiv.frx":007D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNext 
      Height          =   315
      Left            =   4110
      TabIndex        =   3
      ToolTipText     =   "Sair da Tela"
      Top             =   780
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Próxima"
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
      MICON           =   "frmFindAtiv.frx":00EB
      PICN            =   "frmFindAtiv.frx":0107
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdFind 
      Default         =   -1  'True
      Height          =   315
      Left            =   4110
      TabIndex        =   2
      ToolTipText     =   "Sair da Tela"
      Top             =   390
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Localizar"
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
      MICON           =   "frmFindAtiv.frx":0261
      PICN            =   "frmFindAtiv.frx":027D
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
      Caption         =   "Localizar em:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   930
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto a Localizar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmFindAtiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nPos As Integer

Private Sub cmdFind_Click()

Dim X As Integer, s As String, t As String

If Trim$(txtFind.text) = "" Then
    MsgBox "Digite o texto a procurar.", vbCritical, "Atenção"
    txtFind.SetFocus
    Exit Sub
End If

t = UCase$(Trim$(txtFind.text))

With frmAtividade.tvLan
    For X = 1 To .Nodes.Count
        Select Case cmbFind.ListIndex
            Case 0 'tudo
                s = UCase$(.Nodes(X).text)
                If InStr(1, s, t, vbBinaryCompare) > 0 Then
                    .Nodes(X).Selected = True
                    .SetFocus
                    nPos = X
                    Exit For
                End If
            Case 1
                If Left$(.Nodes(X).Key, 6) = "ATIVTL" Then
                    s = UCase$(.Nodes(X).text)
                    If InStr(1, s, t, vbBinaryCompare) > 0 Then
                        .Nodes(X).Selected = True
                        .SetFocus
                        nPos = X
                        Exit For
                    End If
                End If
            Case 2
                If Left$(.Nodes(X).Key, 4) = "ISSF" Or Left$(.Nodes(X).Key, 4) = "ISSV" Or Left$(.Nodes(X).Key, 4) = "ISSE" Then
                    s = UCase$(.Nodes(X).text)
                    If InStr(1, s, t, vbBinaryCompare) > 0 Then
                        .Nodes(X).Selected = True
                        .SetFocus
                        nPos = X
                        Exit For
                    End If
                End If
            Case 3
                If Left$(.Nodes(X).Key, 6) = "VSITEM" Then
                    s = UCase$(.Nodes(X).text)
                    If InStr(1, s, t, vbBinaryCompare) > 0 Then
                        .Nodes(X).Selected = True
                        .SetFocus
                        nPos = X
                        Exit For
                    End If
                End If
        End Select
    Next
End With

End Sub

Private Sub cmdNext_Click()
Dim X As Integer, s As String, t As String

If Trim$(txtFind.text) = "" Then
    MsgBox "Digite o texto a procurar.", vbCritical, "Atenção"
    txtFind.SetFocus
    Exit Sub
End If

t = UCase$(Trim$(txtFind.text))

With frmAtividade.tvLan
    For X = nPos + 1 To .Nodes.Count
        Select Case cmbFind.ListIndex
            Case 0 'tudo
                s = UCase$(.Nodes(X).text)
                If InStr(1, s, t, vbBinaryCompare) > 0 Then
                    .Nodes(X).Selected = True
                    .SetFocus
                    nPos = X
                    Exit For
                End If
            Case 1
                If Left$(.Nodes(X).Key, 6) = "ATIVTL" Then
                    s = UCase$(.Nodes(X).text)
                    If InStr(1, s, t, vbBinaryCompare) > 0 Then
                        .Nodes(X).Selected = True
                        .SetFocus
                        nPos = X
                        Exit For
                    End If
                End If
            Case 2
                If Left$(.Nodes(X).Key, 4) = "ISSF" Or Left$(.Nodes(X).Key, 4) = "ISSV" Or Left$(.Nodes(X).Key, 4) = "ISSE" Then
                    s = UCase$(.Nodes(X).text)
                    If InStr(1, s, t, vbBinaryCompare) > 0 Then
                        .Nodes(X).Selected = True
                        .SetFocus
                        nPos = X
                        Exit For
                    End If
                End If
            Case 3
                If Left$(.Nodes(X).Key, 6) = "VSITEM" Then
                    s = UCase$(.Nodes(X).text)
                    If InStr(1, s, t, vbBinaryCompare) > 0 Then
                        .Nodes(X).Selected = True
                        .SetFocus
                        nPos = X
                        Exit For
                    End If
                End If
        End Select
    Next
End With

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
txtFind.SetFocus
End Sub

Private Sub Form_Load()
Dim X As Long
X = SetParent(Me.hwnd, frmAtividade.hwnd)
cmbFind.ListIndex = 0

End Sub

