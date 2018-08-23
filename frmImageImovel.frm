VERSION 5.00
Begin VB.Form frmImageImovel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foto(s) do Imóvel Selecionado"
   ClientHeight    =   4935
   ClientLeft      =   2625
   ClientTop       =   2115
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7995
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7995
      Begin VB.CommandButton uFoto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Picture         =   "frmImageImovel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Próxima Foto"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton pFoto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Picture         =   "frmImageImovel.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Foto Anterior"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "de"
         Height          =   195
         Left            =   870
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblFotoAte 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   195
         Left            =   1140
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblFotoDe 
         Caption         =   "0"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox txtWait 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1620
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmImageImovel.frx":0294
      Top             =   2340
      Width           =   4215
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2400
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Image Img 
      Height          =   4065
      Left            =   60
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6525
   End
End
Attribute VB_Name = "frmImageImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSubPath As String
Dim nSetor As Integer
Dim nQuadra As Integer
Dim nLote As Integer
Dim nSeq As Integer

Dim pbWidthPercentage As Double
Dim pbHeightPercentage As Double
Dim pbAspectRatio As Double
Private mPics() As StdPicture

Private Sub Form_Activate()
On Error GoTo Erro




If NomeDoComputador = "SKYNET" Then
    Dim mStream As New ADODB.stream
    Dim rst As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    Dim nTotal As Integer
    txtWait.Visible = False
    ConectaBinary
    Sql = "select COUNT(*)AS CONTADOR from F001 where codigo=" & Val(frmCadImob.lblCodReduz.Caption)
    Set RdoAux = cnBinary.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    If RdoAux.RowCount > 0 Then
    nTotal = RdoAux!Contador
    If nTotal = 0 Then
        MsgBox "Não existem fotos para este imóvel."
        Exit Sub
    End If
    
        adoConn.CursorLocation = adUseClient
        adoConn.Open cnBinary.Connect
        
 '       Set FmtPic = New StdFormat.StdDataFormat
 '      FmtPic.Type = fmtPicture
        
        ReDim mPics(0)
        
        rst.Open "Select seq,foto from F001 where codigo=" & Val(frmCadImob.lblCodReduz.Caption), adoConn, adOpenKeyset, adLockOptimistic
        If rst.EOF Then
           MsgBox "Não existem fotos para este imóvel."
        Else
'            Do Until rst.EOF
               

'                Set rst.Fields![Foto].DataFormat = FmtPic
                'img.Picture = rst![Foto].value
                With mStream
                    .Type = adTypeBinary
                    .Open
                    If Not IsNull(rst("foto")) Then
                        .Write rst("foto")
                        img.DataField = "foto"
'                        ReDim Preserve mPics(UBound(mPics) + 1)
'                        mPics(UBound(mPics)) = LoadPicture("foto")
                        Set img.DataSource = rst
                    End If
                End With
                Set mStream = Nothing
                pbWidthPercentage = img.Width / Me.Width
                pbHeightPercentage = img.Height / Me.Height
                pbAspectRatio = img.Height / img.Width
'                rst.MoveNext
'            Loop
        End If

 '   End If
'    RdoAux.Close
    


    Exit Sub
End If




If UCase(NomeDoComputador) = "TESLA" Then
    txtWait.Visible = False
    img.Picture = LoadPicture(sPathBin & "\FotoTeste.jpg")
    Exit Sub
End If

img.Visible = False
File1.Visible = False
txtWait.Visible = True
txtWait.Text = "Aguarde...." & vbCrLf & "Localizando Fotos do Imóvel."

If sFormFoto = "I" Then
    nDist = Left$(frmCadImob.lblDist.Caption, 2)
    nSetor = frmCadImob.lblSetor.Caption
    nQuadra = frmCadImob.txtQuadra.Text
    nLote = frmCadImob.txtLote.Text
    nSeq = frmCadImob.txtSeq.Text
Else
    nDist = frmCadMob.lblDist.Caption
    nSetor = frmCadMob.lblSetor.Caption
    nQuadra = frmCadMob.lblQuadra.Caption
    nLote = frmCadMob.lblLote.Caption
    nSeq = frmCadMob.lblSeq.Caption
End If
If nDist > 1 Then
    txtWait.Text = "Não existem fotos disponíveis para este imóvel."
    Exit Sub
End If
sSubPath = Trim$(sPathFoto)
If nSetor = 1 Then
    sSubPath = sSubPath & "\FOTOS_S1"
ElseIf nSetor = 2 Then
    sSubPath = sSubPath & "\FOTOS_S2"
ElseIf nSetor = 3 Then
    sSubPath = sSubPath & "\FOTOS_S3"
ElseIf nSetor = 4 Then
    sSubPath = sSubPath & "\FOTOS_S4"
End If

'sSubPath = "\\192.168.200.130\fotosgti"
On Error GoTo Erro
If Dir(sSubPath, vbDirectory) = "" Then
   txtWait.Text = "Não existem fotos disponíveis para este imóvel."
   lblFotoDe.Caption = "0"
   lblFotoAte.Caption = "0"
   Frame1.Refresh
   Exit Sub
Else
   
   R = Format(nDist, "00") & "-" & Format(nSetor, "00") & "-" & Format(nQuadra, "0000") & "-" & Format(nLote, "00000") & "*.jpg"
   File1.Pattern = R
   File1.Path = sSubPath
   If File1.ListCount > 0 Then
      txtWait.Text = ""
      txtWait.Visible = False
      img.Visible = True
      File1.ListIndex = 0
      img.Picture = LoadPicture(sSubPath & "\" & File1.FileName)
      lblFotoDe.Caption = "1"
      lblFotoAte.Caption = File1.ListCount
      Frame1.Refresh
 Else
      txtWait.Text = "Não existem fotos disponíveis para este imóvel."
      lblFotoDe.Caption = "0"
      lblFotoAte.Caption = "0"
      Frame1.Refresh
   End If
End If

Exit Sub
Erro:
MsgBox "O diretório de fotos do servidor não está disponível", vbCritical, "Erro"
'Resume Next
Unload Me
End Sub

Private Sub Form_DblClick()
frmZoom.show
frmZoom.ZOrder 0
frmZoom.Left = 500
frmZoom.Top = 500
End Sub

Private Sub Form_Load()
 
'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Form_Resize()

Frame1.Width = Me.Width
img.Width = Me.Width - 50
img.Height = Me.Height - 600
img.Top = 600
img.Left = 0
img.Stretch = True
img.Refresh

'    img.Width = Me.Width * pbWidthPercentage
'    img.Height = img.Width * pbAspectRatio
'    If img.Height > Me.Height Then
'        img.Height = Me.Height * pbHeightPercentage
'        img.Width = img.Height / pbAspectRatio
'    End If
    
End Sub

Private Sub pFoto_Click()

If File1.ListCount = 0 Then Exit Sub

If Val(lblFotoDe.Caption) = 1 Then
   Exit Sub
Else
    On Error Resume Next
   lblFotoDe.Caption = Val(lblFotoDe.Caption) - 1
   lblFotoDe.Refresh
   File1.ListIndex = File1.ListIndex - 1
   Me.MousePointer = vbHourglass
   img.Picture = LoadPicture(sSubPath & "\" & File1.FileName)
   Me.MousePointer = vbDefault
End If

End Sub

Private Sub uFoto_Click()

If File1.ListCount = 0 Then Exit Sub

If Val(lblFotoDe.Caption) = File1.ListCount Then
   Exit Sub
Else
    On Error Resume Next
   lblFotoDe.Caption = File1.ListIndex + 2
   lblFotoDe.Refresh
   File1.ListIndex = File1.ListIndex + 1
   Me.MousePointer = vbHourglass
   img.Picture = LoadPicture(sSubPath & "\" & File1.FileName)
   Me.MousePointer = vbDefault
End If

End Sub

