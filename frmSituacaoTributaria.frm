VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmSituacaoTributaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situação Tributária"
   ClientHeight    =   5325
   ClientLeft      =   4290
   ClientTop       =   2850
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9420
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   45
      TabIndex        =   2
      Top             =   -45
      Width           =   9285
      Begin VB.TextBox txtDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2025
         TabIndex        =   0
         Top             =   215
         Width           =   2040
      End
      Begin prjChameleon.chameleonButton btConsultar 
         Default         =   -1  'True
         Height          =   315
         Left            =   4230
         TabIndex        =   1
         ToolTipText     =   "Consulta Cidadão"
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "C&onsultar"
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
         MICON           =   "frmSituacaoTributaria.frx":0000
         PICN            =   "frmSituacaoTributaria.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Situação..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   6570
         TabIndex        =   6
         Top             =   225
         Width           =   915
      End
      Begin VB.Label lblStatus 
         Caption         =   "NEGATIVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   7605
         TabIndex        =   5
         Top             =   195
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Nº do CPF/CNPJ...:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   225
         Width           =   2130
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdMain 
      Height          =   4650
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   8202
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderDragReorderColumns=   0   'False
      HeaderHotTrack  =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "frmSituacaoTributaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Doc As String

Public Property Let sDoc(sNumDoc As String)
    Doc = sNumDoc
End Property

Private Sub btConsultar_Click()
Dim qd As New rdoQuery, RdoAux As rdoResultset
Dim MaxCod As Integer, sRet As String, Sql As String

lblStatus.Caption = ""

If Len(txtDoc.Text) = 11 Then
    If Not ValidaCPF(txtDoc.Text) Then
        MsgBox "CPF inválido.", vbCritical, "ERRO"
        Exit Sub
    End If
ElseIf Len(txtDoc.Text) = 14 Then
    If Not ValidaCGC(txtDoc.Text) Then
        MsgBox "CNPJ inválido.", vbCritical, "ERRO"
        Exit Sub
    End If
Else
    MsgBox "Nº de Documento inválido.", vbCritical, "ERRO"
    Exit Sub
End If


On Error Resume Next
RdoAux.Close
On Error GoTo 0
Set qd.ActiveConnection = cn
Ocupado
grdMain.Clear
qd.Sql = "{ Call spCDB2(?,?,?,?,?) }"
qd(0) = txtDoc.Text
qd(1) = IIf(Len(txtDoc.Text) = 11, 3, 2)
qd(2) = "35/2007"
qd(3) = "GTI"
qd(4) = 0
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
grdMain.Redraw = False
With RdoAux
    Do Until .EOF
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, !Codigo, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 2, !Nome, DT_LEFT
        sRet = Replace(!resultado, "Certidão ", "")
        grdMain.CellDetails grdMain.Rows, 3, sRet, DT_LEFT
        grdMain.CellDetails grdMain.Rows, 4, !CPF, DT_LEFT
        grdMain.CellDetails grdMain.Rows, 5, !Cnpj, DT_LEFT
           
       .MoveNext
    Loop
   .Close
End With
grdMain.Redraw = True
If grdMain.Rows = 0 Then
    MsgBox "Documento não localizado no cadastro.", vbCritical, "Atenção"
Else
    VerificaStatus
End If
Liberado

End Sub

Private Sub VerificaStatus()

Dim x As Integer, nNeg As Integer, nPos As Integer, nPN As Integer

For x = 1 To grdMain.Rows
    If grdMain.CellText(x, 3) = "Negativa" Then
        nNeg = nNeg + 1
    ElseIf grdMain.CellText(x, 3) = "Positiva" Then
        nPos = nPos + 1
    ElseIf grdMain.CellText(x, 3) = "Positiva com Efeito de Negativa" Then
        nPN = nPN + 1
    End If
Next

If nPN > 0 Or nPos > 0 Then
    lblStatus.Caption = "Positiva"
Else
    lblStatus.Caption = "Negativa"
End If

End Sub

Private Sub Form_Load()
Centraliza Me
lblStatus.Caption = ""
GridHeader
If Doc <> "" Then
    txtDoc.Text = RetornaNumero(Doc)
    btConsultar_Click
End If
sDoc = ""
End Sub

Private Sub txtDoc_Change()
lblStatus.Caption = ""
grdMain.Clear
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
Tweak txtDoc, KeyAscii, IntegerPositive
End Sub

Private Sub GridHeader()
With grdMain
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    .DefaultRowHeight = 17
    .AddColumn "kCod", "Código", ecgHdrTextALignCentre, , 60
    .AddColumn "kNom", "Nome", ecgHdrTextALignLeft, , 250
    .AddColumn "kSit", "Situação", ecgHdrTextALignLeft, , 70
    .AddColumn "kCpf", "CPF", ecgHdrTextALignCentre, , 100
    .AddColumn "kCnp", "CNPJ", ecgHdrTextALignCentre, , 100
End With

End Sub

