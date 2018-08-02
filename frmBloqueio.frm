VERSION 5.00
Begin VB.Form frmBloqueio 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4860
   ClientLeft      =   7125
   ClientTop       =   4545
   ClientWidth     =   9855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   4845
      Left            =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8546
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   255
      ColorTo         =   255
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "O SISTEMA ESTÁ TEMPORARIAMENTE BLOQUEADO. POR FAVOR AGUARDE... O DESBLOQUEIO SERÁ AUTOMÁTICO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   450
         TabIndex        =   0
         Top             =   1500
         Width           =   8985
      End
   End
End
Attribute VB_Name = "frmBloqueio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Centraliza Me
End Sub

