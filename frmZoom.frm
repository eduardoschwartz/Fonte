VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoom da Foto"
   ClientHeight    =   3015
   ClientLeft      =   375
   ClientTop       =   2190
   ClientWidth     =   2775
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   Begin VB.CommandButton cmdMore 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   870
      TabIndex        =   2
      Top             =   2730
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmdLess 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   570
      TabIndex        =   1
      Top             =   2730
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   375
      Top             =   3825
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   0
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   0
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblScale 
      Caption         =   "1x"
      Height          =   210
      Left            =   2445
      TabIndex        =   4
      Top             =   2775
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "Escala Atual:"
      Height          =   210
      Left            =   1320
      TabIndex        =   3
      Top             =   2775
      Width           =   1065
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sHDC As Long
Public sFactor1 As Integer
Public sFactor2 As Integer
Public sFactor3 As Integer

Private Sub cmdLess_Click()
  'Decrease camera image scale factor.
  If sFactor1 > 1 Then
    sFactor1 = sFactor1 - 1
    sFactor2 = 150 / sFactor1
    lblScale.Caption = Trim$(Str(sFactor1)) & "x"
  End If
End Sub

Private Sub cmdMore_Click()
  'Increase camera image scale factor.
  If sFactor1 < 10 Then
    sFactor1 = sFactor1 + 1
    sFactor2 = 150 / sFactor1
    lblScale.Caption = Trim$(Str(sFactor1)) & "x"
  End If
End Sub

Private Sub Form_Activate()
'Limits the Cursor movement to within the form.
Dim client As RECT
Dim upperleft As POINTAPI
  
    'Get information about our wndow
    GetClientRect frmImageImovel.hwnd, client
    upperleft.X = client.Left
    upperleft.Y = client.Top
    'Convert window coördinates to screen coördinates
    ClientToScreen frmImageImovel.hwnd, upperleft
    'move our rectangle
    OffsetRect client, upperleft.X, upperleft.Y
    'limit the cursor movement
    ClipCursor client
  
  sFactor1 = 1
  sFactor2 = 150
  sHDC = GetDC(0)
  frmZoom.Left = 0
  frmZoom.Top = 0
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, -1, 0, 0, 150, 150, &H1 Or &H2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Releases the cursor limits
    ClipCursor ByVal 0&

End Sub

Private Sub Timer1_Timer()
  Dim M As POINTAPI

  'Get cursor position copy screen image and adjust to scale factor.
  GetCursorPos M
  sFactor3 = sFactor2 / 2
  StretchBlt Picture1.hDC, 0, 0, 200, 200, sHDC, M.X - sFactor3, M.Y - sFactor3, sFactor2, sFactor2, vbSrcCopy
  Picture1.Refresh
End Sub
