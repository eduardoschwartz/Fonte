VERSION 5.00
Begin VB.Form frmSenhaMonitor 
   AutoRedraw      =   -1  'True
   Caption         =   "Controle de Senhas - Prefeitura Municipal de Jaboticabal - Gestão de Tributação Municipal Integrada (GTI)"
   ClientHeight    =   5340
   ClientLeft      =   1050
   ClientTop       =   2415
   ClientWidth     =   10755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   10755
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7740
      Top             =   6120
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   330
      Left            =   6750
      TabIndex        =   1
      Top             =   4725
      Width           =   1410
   End
   Begin VB.CommandButton cmdMax 
      Caption         =   "Maximizar"
      Height          =   330
      Left            =   8280
      TabIndex        =   0
      Top             =   4725
      Width           =   1410
   End
   Begin VB.PictureBox PicOld 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H00808080&
      Height          =   2445
      Index           =   2
      Left            =   6930
      ScaleHeight     =   2445
      ScaleWidth      =   2850
      TabIndex        =   5
      Top             =   5670
      Width           =   2850
      Begin VB.Label lblOldG2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   48
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1185
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Top             =   1485
         Width           =   2760
      End
      Begin VB.Label lblOld 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "325"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   65.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1815
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   -450
         Width           =   3435
      End
      Begin VB.Label lblOldG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guiche"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   26.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Index           =   2
         Left            =   315
         TabIndex        =   12
         Top             =   1260
         Width           =   2760
      End
   End
   Begin VB.PictureBox PicOld 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H00808080&
      Height          =   2580
      Index           =   1
      Left            =   6840
      ScaleHeight     =   2580
      ScaleWidth      =   2850
      TabIndex        =   4
      Top             =   2745
      Width           =   2850
      Begin VB.Label lblOldG2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   48
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1185
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   1530
         Width           =   2760
      End
      Begin VB.Label lblOld 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "325"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   65.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1815
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3435
      End
      Begin VB.Label lblOldG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guiche"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   26.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1140
         Index           =   1
         Left            =   315
         TabIndex        =   10
         Top             =   1485
         Width           =   2760
      End
   End
   Begin VB.PictureBox PicOld 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H00808080&
      Height          =   2535
      Index           =   0
      Left            =   6840
      ScaleHeight     =   2535
      ScaleWidth      =   2850
      TabIndex        =   3
      Top             =   0
      Width           =   2850
      Begin VB.Label lblOldG2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   48
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1230
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   1260
         Width           =   2760
      End
      Begin VB.Label lblOldG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guiche"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   26.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1050
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   765
         Width           =   2760
      End
      Begin VB.Label lblOld 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "325"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   65.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1815
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   -720
         Width           =   3435
      End
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00C00000&
      ForeColor       =   &H00808080&
      Height          =   8205
      Left            =   0
      ScaleHeight     =   8205
      ScaleWidth      =   6855
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      Begin VB.Label lblSenha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   60
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   2940
         Left            =   0
         TabIndex        =   18
         Top             =   -135
         Width           =   8745
      End
      Begin VB.Label lblMainG2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   129.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2985
         Left            =   90
         TabIndex        =   14
         Top             =   5175
         Width           =   7395
      End
      Begin VB.Label lblMainG 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guiche"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   80.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   2220
         Left            =   0
         TabIndex        =   7
         Top             =   5310
         Width           =   7395
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "220"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   180
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5820
         Left            =   180
         TabIndex        =   6
         Top             =   1485
         Width           =   8745
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   2430
         Picture         =   "frmSenhaMonitor.frx":0000
         Top             =   2925
         Visible         =   0   'False
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmSenhaMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim soundfile As String

Private Sub cmdMax_Click()
    Me.WindowState = 2
'    Call AlwaysOnTop(Me, True)
    HideMouse
    cmdMax.Visible = False
    cmdSair.Visible = False
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Me.BorderStyle = 2
'    Call AlwaysOnTop(Me, False)
    Me.WindowState = 0
    cmdMax.Visible = True
    cmdSair.Visible = True
    ShowMouse
ElseIf KeyCode = vbKeyF12 Then
    Me.BorderStyle = 0
'    Call AlwaysOnTop(Me, True)
    Me.WindowState = 2
    cmdMax.Visible = False
    cmdSair.Visible = False
    HideMouse
End If
End Sub

Private Sub Form_Load()
soundfile = App.Path & "\Monitor.wav"
ToggleScreenSaverActive False

'Call AlwaysOnTop(Me, True)
cmdMax_Click
HideMouse
Limpa
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ShowMouse
ToggleScreenSaverActive False

End Sub

Private Sub Form_Resize()
Dim x As Integer

cmdMax.Left = Me.Width - cmdMax.Width - 500
cmdSair.Left = Me.Width - cmdSair.Width - 2000
cmdMax.Top = Me.Height - cmdMax.Height - 800
cmdSair.Top = Me.Height - cmdSair.Height - 800

With picMain
    .Width = Me.Width / 2 + Me.Width / 4
    .Height = Me.Height
End With

With PicOld(0)
    .Left = picMain.Width
    .Width = Me.Width / 4
    .Height = Me.Height / 3
End With
With PicOld(1)
    .Left = picMain.Width
    .Width = Me.Width / 4
    .Height = Me.Height / 3
    .Top = PicOld(0).Height
End With
With PicOld(2)
    .Left = picMain.Width
    .Width = Me.Width / 4
    .Height = Me.Height / 3
    .Top = PicOld(0).Height * 2
End With
lblMain.Width = picMain.Width
lblMain.Left = picMain.Left
lblMain.Top = 400
lblMainG.Width = picMain.Width
lblMainG.Left = picMain.Left
lblMainG.Top = lblMain.Top + lblMain.Height - 1900
lblMainG2.Width = picMain.Width
lblMainG2.Left = picMain.Left
lblMainG2.Top = lblMain.Top + lblMain.Height - 700
For x = 0 To 2
    lblOld(x).Width = PicOld(x).Width
    lblOld(x).Left = PicOld(x).Left - picMain.Width
    lblOldG(x).Width = PicOld(x).Width
    lblOldG(x).Left = PicOld(x).Left - picMain.Width
Next

lblOld(0).Top = -100
lblOld(1).Top = -100
lblOld(2).Top = -300
lblOldG(0).Top = lblOld(0).Top + lblOld(0).Height - 300
lblOldG(1).Top = lblOld(1).Top + lblOld(1).Height - 300
lblOldG(2).Top = lblOld(2).Top + lblOld(2).Height - 400
lblOldG2(0).Top = lblOld(0).Top + lblOld(0).Height
lblOldG2(1).Top = lblOld(1).Top + lblOld(1).Height
lblOldG2(2).Top = lblOld(2).Top + lblOld(2).Height - 100

UpdateBltSample
End Sub

Private Sub picMain_Resize()
picMain.Cls
End Sub

Private Sub PicOld_Resize(Index As Integer)
PicOld(Index).Cls
End Sub

Private Sub UpdateBltSample()

Dim PicRect As RECT
Dim xyOffset As Long, Index As Integer
xyOffset = 0

PicRect.Left = xyOffset
PicRect.Top = xyOffset
PicRect.Right = picMain.ScaleWidth / Screen.TwipsPerPixelX - xyOffset
PicRect.Bottom = picMain.ScaleHeight / Screen.TwipsPerPixelY - xyOffset

TileBltRectEx picMain.hdc, PicRect, Image1.Picture, 1, 1
picMain.Line (picMain.Left + 50, picMain.Top + 50)-(picMain.Width - 50, picMain.Height - 50), , B

For Index = 0 To 2
    TileBltRectEx PicOld(Index).hdc, PicRect, Image1.Picture, 1, 1
    PicOld(Index).Line (50, 50)-(PicOld(Index).Width - 50, PicOld(Index).Height - 50), , B
Next

End Sub

Private Sub Limpa()
Dim x As Integer
lblMain.Caption = ""
lblMainG.Caption = ""
lblMainG2.Caption = ""
For x = 0 To 2
    lblOld(x).Caption = ""
    lblOldG(x).Caption = ""
    lblOldG2(x).Caption = ""
Next
End Sub

Private Sub HideMouse()
'Do While ShowCursor(False) >= 0
'DoEvents
'Loop
End Sub

Private Sub ShowMouse()
'Do While ShowCursor(True) < 0
'Loop
End Sub

Private Sub Timer1_Timer()
Dim RdoAux As rdoResultset, Sql As String, nSenha As Integer
Dim returnval As Long
nSenha = 0
If cGetInputState() <> 0 Then Me.Refresh
Sql = "SELECT * FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND DATACHAMADA IS NOT NULL AND MONITOR=0 "
Sql = Sql & "ORDER BY HORACHAMADA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then Exit Sub
    
    nSenha = !SENHA
    lblOld(0).Caption = lblOld(1).Caption
    lblOldG(0).Caption = lblOldG(1).Caption
    lblOldG2(0).Caption = lblOldG2(1).Caption
    lblOld(1).Caption = lblOld(2).Caption
    lblOldG(1).Caption = lblOldG(2).Caption
    lblOldG2(1).Caption = lblOldG2(2).Caption
    lblOld(2).Caption = lblMain.Caption
    lblOldG(2).Caption = lblMainG.Caption
    lblOldG2(2).Caption = lblMainG2.Caption
    lblMain.Caption = Format(nSenha, "000")
    lblMainG.Caption = "Guiche"
    lblMainG2.Caption = Format(!GUICHE, "00")
    
    Me.Refresh
    .Close
End With

Sql = "UPDATE SSPAC SET MONITOR=1 "
Sql = Sql & " WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now)
Sql = Sql & " AND DAY(DATAENTRADA)=" & Day(Now) & " AND SENHA=" & nSenha
cn.Execute Sql, rdExecDirect
If nSenha > 0 Then
    returnval = PlaySound(soundfile, 0, &H0)
End If
End Sub
