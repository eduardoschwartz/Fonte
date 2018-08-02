VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmOficio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar/Receber Ofícios"
   ClientHeight    =   6000
   ClientLeft      =   2895
   ClientTop       =   3780
   ClientWidth     =   11325
   ForeColor       =   &H00000000&
   Icon            =   "frmOficio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   11325
   Begin Tributacao.jcFrames frEnviado 
      Height          =   5415
      Left            =   4230
      Top             =   45
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   9551
      FillColor       =   16777215
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Ofícios enviados/Enviar novo Ofício"
      TextColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   16777215
      ColorTo         =   14737632
   End
   Begin Tributacao.jcFrames frBotao 
      Height          =   465
      Left            =   4230
      Top             =   5490
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   820
      FrameColor      =   8421504
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin prjChameleon.chameleonButton cmdAlterar 
         Height          =   315
         Left            =   1665
         TabIndex        =   3
         ToolTipText     =   "Editar Registro"
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Editar"
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
         MICON           =   "frmOficio.frx":01CA
         PICN            =   "frmOficio.frx":01E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdPrint 
         Height          =   315
         Left            =   4185
         TabIndex        =   1
         ToolTipText     =   "Imprime o Carnê de Parcelamento"
         Top             =   75
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Imprimir"
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
         MICON           =   "frmOficio.frx":0340
         PICN            =   "frmOficio.frx":035C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNovo 
         Height          =   315
         Left            =   405
         TabIndex        =   2
         ToolTipText     =   "Novo Registro"
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "frmOficio.frx":04B6
         PICN            =   "frmOficio.frx":04D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdExcluir 
         Height          =   315
         Left            =   2925
         TabIndex        =   4
         ToolTipText     =   "Excluir Registro"
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "E&xcluir"
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
         MCOL            =   16776960
         MPTR            =   1
         MICON           =   "frmOficio.frx":062C
         PICN            =   "frmOficio.frx":0648
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAnexar 
         Height          =   315
         Left            =   5445
         TabIndex        =   5
         ToolTipText     =   "Anexar um Processo"
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "A&nexos"
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
         MICON           =   "frmOficio.frx":06EA
         PICN            =   "frmOficio.frx":0706
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin Tributacao.jcFrames frRecebido 
      Height          =   5415
      Left            =   4230
      Top             =   45
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   9551
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Ofícios Recebidos"
      TextColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   14737632
      ColorTo         =   16777215
   End
   Begin Tributacao.jcFrames jcFrames 
      Height          =   5910
      Index           =   1
      Left            =   45
      Top             =   45
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   10425
      FillColor       =   16777215
      TextBoxColor    =   11595760
      Style           =   3
      RoundedCorner   =   0   'False
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Lista de Ofícios"
      TextColor       =   8388608
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSComctlLib.TreeView tvMain 
         Height          =   5535
         Left            =   45
         TabIndex        =   0
         Top             =   360
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   9763
         _Version        =   393217
         Indentation     =   794
         LabelEdit       =   1
         Style           =   6
         HotTracking     =   -1  'True
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmOficio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const WM_HSCROLL As Long = &H114
Private Const SB_LINELEFT As Long = 0

Private Sub cmdNovo_Click()
frmOficioInfo.show 1
End Sub

Private Sub Form_Load()
Centraliza Me

Buildtree

End Sub

Private Sub Buildtree()
Dim Sql As String, RdoAux As rdoResultset, x As Integer, i As Long
With tvMain
    
    Set NodX = .Nodes.Add(, , "REC", "Recebidos")
    Set NodX = .Nodes.Add(, , "ENV", "Enviados")
    
    Sql = "SELECT centrocusto.CODIGO, centrocusto.DESCRICAO, usuariocc.NOME FROM centrocusto INNER JOIN usuariocc ON centrocusto.CODIGO = usuariocc.CODIGOCC "
    Sql = Sql & "WHERE centrocusto.ATIVO=1 AND usuariocc.NOME='" & NomeDeLogin & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Set NodX = tvMain.Nodes.Add("REC", tvwChild, "REC" & Format(!Codigo, "000"), !Descricao)
            tvMain.Nodes("REC" & Format(!Codigo, "000")).ForeColor = vbBlue
           .MoveNext
        Loop
       .Close
    End With
End With

With tvMain
    For x = 1 To .Nodes.Count
       .Nodes(x).EnsureVisible
    Next
   .Nodes("ENV").Bold = True
   .Nodes("REC").Bold = True
   .Nodes("ENV").ForeColor = Roxo
   .Nodes("REC").ForeColor = &H4040&
End With
 
'For i = 1 To 4 'Scroll left 4 times to bring the icon into view
'    SendMessage tvMain.hwnd, WM_HSCROLL, SB_LINELEFT, 0&
'Next i
  
tvMain.Nodes("REC").EnsureVisible
tvMain.Nodes("REC").Selected = True
frEnviado.Visible = False
frRecebido.Visible = True
    
End Sub

Private Sub tvMain_NodeClick(ByVal Node As MSComctlLib.Node)

If Left(Node.Key, 3) = "REC" Then
    frEnviado.Visible = False
    frRecebido.Visible = True
Else
    frEnviado.Visible = True
    frRecebido.Visible = False
End If

End Sub
