VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmSerasa 
   Caption         =   "Consulta ao Serasa"
   ClientHeight    =   6360
   ClientLeft      =   2955
   ClientTop       =   1965
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9645
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      ExtentX         =   16933
      ExtentY         =   11139
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmSerasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Centraliza Me
Form_Resize
WebBrowser1.Navigate "www.serasaexperian.com.br"
End Sub

Private Sub Form_Resize()
With WebBrowser1
'    .Top = Me.Top - 1200
'    .Left = 100
    .Width = Me.Width - 200
    .Height = Me.Height - 300
End With
End Sub
