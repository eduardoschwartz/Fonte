VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCnsRua 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Imóveis no Logradouro"
   ClientHeight    =   6255
   ClientLeft      =   1485
   ClientTop       =   1725
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   5970
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   0
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   655
      TabIndex        =   2
      Top             =   0
      Width           =   9885
      Begin VB.PictureBox Pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2235
         Left            =   -30
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3333
         TabIndex        =   3
         Top             =   -30
         Width           =   50000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   435
         Left            =   -30
         TabIndex        =   4
         Top             =   2160
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   767
         _Version        =   393216
         SmallChange     =   3
         TickStyle       =   1
         TickFrequency   =   3
      End
   End
   Begin VB.PictureBox picContainer 
      Height          =   345
      Left            =   5040
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   0
      Top             =   1170
      Visible         =   0   'False
      Width           =   225
      Begin VB.PictureBox picSize 
         Height          =   960
         Left            =   45
         ScaleHeight     =   900
         ScaleWidth      =   2865
         TabIndex        =   1
         Top             =   45
         Width           =   2925
      End
   End
   Begin VB.Image img 
      Height          =   3435
      Left            =   5580
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   4005
   End
   Begin VB.Label Label1 
      Caption         =   "Código Reduzido:"
      Height          =   225
      Left            =   30
      TabIndex        =   16
      Top             =   2850
      Width           =   1275
   End
   Begin VB.Label lblCodReduz 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1440
      TabIndex        =   15
      Top             =   2850
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Bairro..................:"
      Height          =   195
      Left            =   30
      TabIndex        =   14
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label lblbairro 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   3240
      Width           =   3945
   End
   Begin VB.Label Label3 
      Caption         =   "Nº de Inscrição..:"
      Height          =   195
      Left            =   30
      TabIndex        =   12
      Top             =   3630
      Width           =   1245
   End
   Begin VB.Label lblIC 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   3630
      Width           =   3945
   End
   Begin VB.Label Label4 
      Caption         =   "Proprietário.........:"
      Height          =   195
      Left            =   30
      TabIndex        =   10
      Top             =   4020
      Width           =   1245
   End
   Begin VB.Label lblProp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   4020
      Width           =   3945
   End
   Begin VB.Label Label5 
      Caption         =   "Quadras.............:"
      Height          =   195
      Left            =   30
      TabIndex        =   8
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label Label6 
      Caption         =   "Lotes.................:"
      Height          =   195
      Left            =   30
      TabIndex        =   7
      Top             =   4890
      Width           =   1245
   End
   Begin VB.Label lblQuadras 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   4440
      Width           =   3945
   End
   Begin VB.Label lblLotes 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   4890
      Width           =   3945
   End
End
Attribute VB_Name = "frmCnsRua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DEFINIÇÃO DAS API'S
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'TIPOS
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Pontos
    Numero As Integer
    X1 As Integer
    Y1 As Integer
End Type

'VARIAVEIS
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim NumPontos As Integer, xPonto As Integer
Dim aPontos() As Pontos, bExec As Boolean, sNomeLog As String
Dim X1 As Integer, Y1 As Integer, nX1 As Integer, nX2 As Integer, nNumero As Integer
Dim mRGN() As Long, R As RECT, X As Long, Y As Long, nRaio As Integer, nMaxX As Integer
Dim nDist As Integer, nSetor As Integer, nQuadra As Integer, nLote As Integer, nSeq As Integer
Dim nNum As Integer

'ALIMENTA MATRIZ
Private Sub SetPoint(X1 As Integer, Y1 As Integer, Num As Integer)

ReDim Preserve aPontos(UBound(aPontos) + 1)
aPontos(UBound(aPontos)).Numero = Num
aPontos(UBound(aPontos)).X1 = X1
aPontos(UBound(aPontos)).Y1 = Y1
If X1 > nMaxX Then
   nMaxX = X1
End If

End Sub

'DEFINE AS REGIÕES
Private Sub MontaMatriz()
Dim X As Integer

Do While Not RdoAux.EOF
    X = RdoAux.AbsolutePosition
    nNumero = RdoAux!Li_Num
    If nNumero Mod 2 <> 0 Then
        X1 = nX1 + 30 + (nRaio * 2)
        Y1 = 50
        nX1 = nX1 + 30
    Else
        X1 = nX2 + 30 + (nRaio * 2)
        Y1 = 100
        nX2 = nX2 + 30
    End If
    SetPoint X1, Y1, nNumero
    SetRect R, aPontos(UBound(aPontos)).X1 - nRaio, aPontos(UBound(aPontos)).Y1 - nRaio, aPontos(UBound(aPontos)).X1 + nRaio, aPontos(UBound(aPontos)).Y1 + nRaio
    ReDim Preserve mRGN(UBound(mRGN) + 1)
    mRGN(UBound(mRGN)) = CreateEllipticRgnIndirect(R)
    RdoAux.MoveNext
Loop

'HScroll1.max = nMaxX - 100
Slider1.Max = nMaxX - 100
End Sub

'MONTA O DESENHO
Private Sub Desenha()

Pic.ScaleMode = 3

Pic.CurrentX = 30 + (Slider1.Value)
Pic.CurrentY = 70
Pic.Print sNomeLog

'PONTO DE INICIO
Pic.FillStyle = 0
Pic.FillColor = vbYellow
Pic.Circle (27, 55), 8, vbBlack
Pic.Circle (27, 95), 8, vbBlack
Pic.FillColor = vbWhite
    'DESENHA AS LINHAS
Pic.Line (35, 55)-(nMaxX + 30, 55), vbBlue
Pic.Line (35, 95)-(nMaxX + 30, 95), vbBlue
For X = 0 To UBound(aPontos)
    If X > 0 Then
        'TEXTO
        Pic.CurrentX = aPontos(X).X1 - nRaio
        If aPontos(X).Numero Mod 2 <> 0 Then
           Pic.CurrentY = aPontos(X).Y1 - 25
        Else
           Pic.CurrentY = 110
        End If
        Pic.Print aPontos(X).Numero
        If X = xPonto Then
            'PONTO CORRENTE
            Pic.FillColor = vbRed
            Pic.Circle (aPontos(X).X1, aPontos(X).Y1), nRaio, vbBlue
        Else
            'PONTO DESATIVADO
            Pic.FillColor = VerdeAccess
            Pic.Circle (aPontos(X).X1, aPontos(X).Y1), nRaio, AmareloClaro
        End If
    End If
Next
'PONTO FINAL
Pic.FillColor = vbYellow
Pic.Circle (nMaxX + 38, 55), 8, vbBlack
Pic.Circle (nMaxX + 38, 95), 8, vbBlack

End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'VERIFICA SE O MOUSE ESTA EM ALGUM PONTO
For z = 0 To UBound(mRGN)
    If PtInRegion(mRGN(z), X, Y) <> 0 Then
        xPonto = z
        'MUDA O CURSOR
        Pic.MouseIcon = LoadResPicture(101, 2)
        Pic.MousePointer = vbCustom
        CarregaImovel
        Exit For
    End If
Next

End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim z As Integer, Achou As Boolean
On Error Resume Next
If cGetInputState() <> 0 Then DoEvents

'ATUALIZA O PONTEIRO DO MOUSE
xPonto = 0
Achou = False
For z = 0 To UBound(mRGN)
  If PtInRegion(mRGN(z), X, Y) <> 0 Then
    xPonto = z
    Desenha
    Pic.MouseIcon = LoadResPicture(101, 2)
    Pic.MousePointer = vbCustom
    Achou = True
    Exit For
  End If
Next

If Not Achou Then Pic.MousePointer = vbDefault

End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DESATIVA O PONTO
xPonto = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'REMOVE DA MEMÓRIA
For X = 0 To UBound(mRGN)
  DeleteObject mRGN(X)
Next
ReDim aPontos(0)
Pic.Cls

End Sub

Private Sub Form_Load()

'ZERA AS VARIAVEIS
ReDim aPontos(0)
ReDim mRGN(0)
bExec = False
bLado = True

Centraliza Me
'INICIALIZA PONTOS

nRaio = 6 'RAIO DOS PONTOS
aPontos(0).X1 = 30 'PRIMEIRO PONTO X
aPontos(0).Y1 = 50 'PRIMEIRO PONTO Y
nX1 = 30
nX2 = 30
Y2 = 100

'Carrega Numeros da Rua
Sql = "SELECT LI_NUM FROM VWCNSNUMIMOVEL WHERE CODLOGR=" & Val(frmLogradouro.txtCod.Text) & " ORDER BY LI_NUM"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    NumPontos = .RowCount 'NUMERO DE PONTOS
   'DESENHO INICIAL
    If Val(NumPontos) > 0 Then
       Sql = "SELECT ABREVTIPOLOG, ABREVTITLOG, NOMELOGRADOURO, CODLOGR "
       Sql = Sql & "FROM vwFACEQUADRA WHERE CODLOGR=" & Val(frmLogradouro.txtCod.Text)
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       sNomeLog = Format(RdoAux2!CodLogr, "0000") & " - " & Trim$(RdoAux2!AbrevTipoLog) & " " & Trim$(SubNull(RdoAux2!AbrevTitLog)) & " " & RdoAux2!NomeLogradouro
       RdoAux2.Close
       Pic.Cls
       MontaMatriz
       Desenha
    End If
   .Close
End With
End Sub

Private Sub CarregaImovel()
Dim sSubPath As String

Ocupado
nNum = aPontos(xPonto).Numero
Sql = "SELECT CODREDUZIDO,DESCBAIRRO,DISTRITO,SETOR,QUADRA,LOTE,SEQ,LI_QUADRAS,LI_LOTES "
Sql = Sql & "FROM VWFULLIMOVEL2 WHERE CODLOGR=" & Val(frmLogradouro.txtCod.Text) & " AND LI_NUM=" & nNum
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblCodReduz.Caption = !CODREDUZIDO
    nDist = !Distrito
    nSetor = !Setor
    nQuadra = !Quadra
    nLote = !Lote
    nSeq = !Seq
    lblBairro.Caption = SubNull(!DescBairro)
    lblQuadras.Caption = SubNull(!Li_Quadras)
    lblLotes.Caption = SubNull(!Li_Lotes)
    lblIC.Caption = nDist & "." & Format(nSetor, "00") & "." & Format(nQuadra, "0000") & "." & Format(nLote, "00000") & "." & Format(nSeq, "00")
    Sql = "SELECT  NOMECIDADAO FROM PROPRIETARIO INNER JOIN CIDADAO ON  PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
    Sql = Sql & "WHERE PROPRIETARIO.CODREDUZIDO =" & !CODREDUZIDO & " AND PROPRIETARIO.TIPOPROP = 'P' AND PROPRIETARIO.PRINCIPAL = 1"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
        lblProp.Caption = RdoAux2!NOMECIDADAO
    End If
    RdoAux2.Close
    sSubPath = sPathFoto
    If nSetor = 1 Then
        sSubPath = sSubPath & "\FOTOS_S1"
    ElseIf nSetor = 2 Then
        sSubPath = sSubPath & "\FOTOS_S2"
    ElseIf nSetor = 3 Then
        sSubPath = sSubPath & "\FOTOS_S3"
    ElseIf nSetor = 4 Then
        sSubPath = sSubPath & "\FOTOS_S4"
    End If
    On Error Resume Next
    If Dir(sSubPath, vbDirectory) = "" Then
       img.Visible = False
       Liberado
       Exit Sub
    Else
       Rss = Format(nDist, "00") & "-" & Format(nSetor, "00") & "-" & Format(nQuadra, "0000") & "-" & Format(nLote, "00000") & "*.jpg"
       File1.Pattern = Rss
       File1.Path = sSubPath
       If File1.ListCount > 0 Then
          img.Visible = True
          File1.ListIndex = 0
          img.Picture = LoadPicture(sSubPath & "\" & File1.FileName)
     Else
          img.Visible = False
       End If
    End If

End With
Liberado

End Sub


Private Sub Slider1_Click()
    Pic.Cls
    Pic.Left = 0 - Slider1.Value
    Desenha
End Sub
