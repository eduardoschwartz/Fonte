VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTributoArtigo 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tributo - Artigo"
   ClientHeight    =   6000
   ClientLeft      =   2475
   ClientTop       =   2670
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8280
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   2145
      Left            =   0
      TabIndex        =   6
      Top             =   3390
      Width           =   8265
      Begin VB.TextBox txtArtigo 
         Appearance      =   0  'Flat
         Height          =   1230
         Left            =   45
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   855
         Width           =   8115
      End
      Begin VB.Label lblTributo 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do Artigo"
         Height          =   195
         Left            =   1395
         TabIndex        =   9
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do Artigo"
         Height          =   195
         Index           =   11
         Left            =   105
         TabIndex        =   8
         Top             =   585
         Width           =   1635
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tributo..............:"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   255
         Width           =   1185
      End
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   7155
      TabIndex        =   5
      ToolTipText     =   "Sair da Tela"
      Top             =   5595
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmTributoArtigo.frx":0000
      PICN            =   "frmTributoArtigo.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTributo 
      Height          =   3360
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5927
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmTributoArtigo.frx":008A
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   7140
      TabIndex        =   4
      ToolTipText     =   "Cancelar Edição"
      Top             =   5595
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmTributoArtigo.frx":011B
      PICN            =   "frmTributoArtigo.frx":0137
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   6060
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   5595
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmTributoArtigo.frx":0291
      PICN            =   "frmTributoArtigo.frx":02AD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   6045
      TabIndex        =   3
      ToolTipText     =   "Gravar os Dados"
      Top             =   5595
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmTributoArtigo.frx":0407
      PICN            =   "frmTributoArtigo.frx":0423
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
Attribute VB_Name = "frmTributoArtigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAlterar_Click()
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdGravar_Click()
Dim MaxCod As Integer, Sql As String, RdoAux As rdoResultset

MaxCod = Val(Left(lblTributo.Caption, 3))
If Trim(txtArtigo.Text) <> "" Then
    Sql = "SELECT * FROM TRIBUTOARTIGO WHERE CODTRIBUTO=" & MaxCod
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
    With RdoAux
        If .RowCount = 0 Then
            Sql = "INSERT TRIBUTOARTIGO (CODTRIBUTO,ARTIGO) VALUES(" & MaxCod & ",'" & Mask(txtArtigo.Text) & "')"
        Else
            Sql = "UPDATE TRIBUTOARTIGO SET ARTIGO='" & Mask(txtArtigo.Text) & "' WHERE CODTRIBUTO=" & MaxCod
        End If
       .Close
    End With
Else
    Sql = "DELETE FROM TRIBUTOARTIGO WHERE CODTRIBUTO=" & MaxCod
End If
cn.Execute Sql, rdExecDirect
grdTributo.TextMatrix(grdTributo.Row, 2) = txtArtigo.Text
    
Eventos "INICIAR"
Evento = ""
End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
grdTributo.Rows = 1
CarregaLista
Eventos "INICIAR"
End Sub

Private Sub CarregaLista()
Dim RdoAux As rdoResultset, Sql As String

Sql = "SELECT tributo.codtributo, tributo.abrevtributo, tributoartigo.artigo FROM tributo LEFT OUTER JOIN "
Sql = Sql & "tributoartigo ON tributo.codtributo = tributoartigo.codtributo ORDER BY tributo.codtributo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

grdTributo.Rows = 1
With RdoAux
    Do Until .EOF
       grdTributo.AddItem Format(!CodTributo, "000") & Chr(9) & !ABREVTRIBUTO & Chr(9) & SubNull(!ARTIGO)
      .MoveNext
    Loop
   .Close
End With
grdTributo_RowColChange
End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdAlterar.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   txtArtigo.BackColor = Kde
   txtArtigo.Locked = True
   grdTributo.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdAlterar.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   grdTributo.Enabled = False
   txtArtigo.BackColor = Branco
   txtArtigo.Locked = False
   txtArtigo.SetFocus
End If

End Sub

Private Sub grdTributo_RowColChange()
With grdTributo
    If .Row > 0 Then
        lblTributo.Caption = .TextMatrix(.Row, 0) & " - " & .TextMatrix(.Row, 1)
        txtArtigo.Text = .TextMatrix(.Row, 2)
    End If
End With

End Sub
