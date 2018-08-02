Attribute VB_Name = "mdlSkin"

Private sFillStyle As sbFillStyle
Private sbFontName As String
Private sbText As String
Private sbForeColor As Long
Private sbBackColor As Long
Private sbFontSize As Integer
Private sbActive As Boolean
Private sAlign As sbAlign
Private Const sBarWidth = 20

Public Property Get FillStyle() As sbFillStyle
 FillStyle = sFillStyle
End Property
Public Property Let FillStyle(sBarFillStyle As sbFillStyle)
 sFillStyle = sBarFillStyle
 'Refresh
End Property
Public Property Get FontName() As String
 FontName = sbFontName
End Property
Public Property Let FontName(sBarFontName As String)
 sbFontName = sBarFontName
 'Refresh
End Property
Public Property Get text() As String
 text = sbText
End Property
Public Property Let text(sBarText As String)
 sbText = sBarText
' Refresh
End Property
Public Property Get BackColor() As Long
 BackColor = sbBackColor
End Property
Public Property Let BackColor(sBarBackColor As Long)
 sbBackColor = sBarBackColor
 'Refresh
End Property
Public Property Get ForeColor() As Long
 ForeColor = sbForeColor
End Property
Public Property Let ForeColor(sBarForeColor As Long)
 sbForeColor = sBarForeColor
' Refresh
End Property
Public Property Get FontSize() As Integer
 FontSize = sbFontSize
End Property
Public Property Let FontSize(sBarFontSize As Integer)
 sbFontSize = sBarFontSize
 'Refresh
End Property
Public Property Get Align() As sbAlign
 Align = sAlign
End Property
Public Property Let Align(sBarAlign As sbAlign)
 sAlign = sBarAlign
 'Refresh
End Property
Public Sub Create(frmDest As Form, sBarText As String, sBarFillStyle As sbFillStyle, sBarAlign As sbAlign, sBarForeColor As Long, sBarBackColor As Long, sBarFontName As String, sBarFontSize As Integer)
 sbText = sBarText: sFillStyle = sBarFillStyle
 sAlign = sBarAlign: sbForeColor = sBarForeColor
 sbBackColor = sBarBackColor: sbFontSize = sBarFontSize
 sbFontName = sBarFontName
 sbActive = True
 Refresh frmDest
End Sub
Public Sub Destroy(frmDest As Form)
 oldMode% = frmDest.ScaleMode
 frmDest.ScaleMode = 3
 frmDest.Line (0, 0)-(sBarWidth + 10, frmDest.Height), frmDest.BackColor, BF
 frmDest.Line (frmDest.ScaleWidth - sBarWidth + 1, 0)-(frmDest.ScaleWidth, frmDest.Height), frmDest.BackColor, BF
 frmDest.ScaleMode = oldMode%
 sbActive = False
End Sub
Private Sub Farbverlauf(col1&, frmDest As Form)
 Const intBLUESTART% = 255
 Const intBLUEEND% = 0
 Const intBANDHEIGHT% = 2
 Dim sngBlueCur As Single
 Dim sngBlueStep As Single
 Dim intFormHeight As Integer
 Dim intFormWidth As Integer
 Dim intY As Integer
 intFormHeight = frmDest.ScaleHeight
 If intFormHeight < 1 Then Exit Sub
 intFormWidth = sBarWidth
 sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
 sngBlueCur = intBLUESTART
 For intY = 0 To intFormHeight Step intBANDHEIGHT
  Select Case col1&
   Case QBColor(1)
       If sAlign = 1 Then
        'frmDest.Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
          frmDest.Line (3, intY - 1)-(intFormWidth + 2, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
     End If
   Case Else
    frmDest.Line (0, 0)-(sBarWidth, frmDest.Height), col1&, BF
  End Select
  sngBlueCur = sngBlueCur + sngBlueStep
 Next intY
 
End Sub
Public Sub AutoResize()
 'Refresh Me
End Sub
Public Sub Refresh(frmDest As Form)
 Destroy frmDest
 sbActive = True
 oldMode% = frmDest.ScaleMode
 frmDest.ScaleMode = 3
 If sbFontName = "" Then sbFontName = "Arial"
 If sFillStyle = 0 Then sFillStyle = sbFilled
 If sAlign = 0 Then sAlign = sbAlignLeft
 If sbFontSize = 0 Then sbFontSize = 14
 frmDest.AutoRedraw = True
 If sFillStyle = sbFilled Then
  If sAlign = 1 Then
   frmDest.Line (0, 0)-(sBarWidth, frmDest.Height), sbBackColor, BF
  Else
   frmDest.Line (frmDest.ScaleWidth - sBarWidth, 0)-(frmDest.ScaleWidth, frmDest.Height), sbBackColor, BF
  End If
 ElseIf sFillStyle = sbSmooth Then
  Farbverlauf sbBackColor, frmDest
 End If
 BckupCol& = frmDest.ForeColor
 frmDest.ForeColor = sbForeColor
 Dim hFont&, fontOld&
 hFont = CreateFont(-sbFontSize, 0, 90 * 10, 0, 700, False, False, 0, 1, 4, &H10, 2, 4, sbFontName)
 fontOld = SelectObject(frmDest.hDC, hFont)
 frmDest.ScaleMode = 3
 If sAlign = 1 Then
  TextOut frmDest.hDC, 4, frmDest.ScaleHeight - 8, sbText, Len(sbText)
 Else
  TextOut frmDest.hDC, (frmDest.ScaleWidth - sBarWidth) + 2, frmDest.ScaleHeight - 8, sbText, Len(sbText)
 End If
 SelectObject frmDest.hDC, fontOld
 DeleteObject hFont
 frmDest.ForeColor = BckupCol&
 frmDest.ScaleMode = oldMode%
End Sub

Public Sub SkinForm(frmDest As Form)

Dim bMinimize As Boolean, bMaximize As Boolean, bClose As Boolean
Dim CTL As Control
bMinimize = frmDest.MinButton
bMaximize = frmDest.MaxButton
bClose = frmDest.ControlBox

'*** ALINHA CONTROLES ****
On Error Resume Next
For X = 0 To frmDest.Controls.Count - 1
      If frmDest.Controls(X).Container.Name = frmDest.Name Then
          frmDest.Controls(X).Top = frmDest.Controls(X).Top + frmDest.ImgSkin(0).Height
          frmDest.Controls(X).Left = frmDest.Controls(X).Left + 370
      End If
Next
On Error GoTo 0

For Each CTL In frmDest
      If TypeOf CTL Is Line Then
           If CTL.Container.Name = frmDest.Name Then
                CTL.X1 = CTL.X1 + 350
                CTL.X2 = CTL.X2 + 350
                CTL.Y1 = CTL.Y1 + 350
                CTL.Y2 = CTL.Y2 + 350
           End If
      End If
Next

frmDest.Height = frmDest.Height + frmDest.ImgSkin(0).Height
frmDest.Top = frmDest.Top - frmDest.ImgSkin(0).Height
frmDest.Width = frmDest.Width + 370
frmDest.Left = frmDest.Left - 200


If Dir(ArqBinImgTmp) <> "" Then Kill ArqBinImgTmp

'*** BARRA DE TITULO ****
frmDest.ImgSkin(0).Height = 330
frmDest.ImgSkin(0).Width = frmDest.Width
gtiObj.LikroTmuna ArqBinImg, "SKTT", ArqBinImgTmp
frmDest.ImgSkin(0).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(0).Picture = LoadPicture(sPathBin & "\SKTT.BMP")
frmDest.ImgSkin(0).Left = 0
frmDest.ImgSkin(0).Top = 0

'*** BORDA ESQUERDA ****
Load frmDest.ImgSkin(1)
frmDest.ImgSkin(1).Height = frmDest.Height - frmDest.ImgSkin(0).Height '
frmDest.ImgSkin(1).Width = 60
gtiObj.LikroTmuna ArqBinImg, "SKLF", ArqBinImgTmp
frmDest.ImgSkin(1).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(1).Picture = LoadPicture(sPathBin & "\SKLF.BMP")
frmDest.ImgSkin(1).Left = -20
frmDest.ImgSkin(1).Top = frmDest.ImgSkin(0).Height
frmDest.ImgSkin(1).Visible = True

'*** BORDA DIREITA ****
Load frmDest.ImgSkin(2)
frmDest.ImgSkin(2).Height = frmDest.Height - frmDest.ImgSkin(0).Height '
frmDest.ImgSkin(2).Width = 80
gtiObj.LikroTmuna ArqBinImg, "SKRG", ArqBinImgTmp
frmDest.ImgSkin(2).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(2).Picture = LoadPicture(sPathBin & "\SKRG.BMP")
frmDest.ImgSkin(2).Left = frmDest.Width - 80
frmDest.ImgSkin(2).Top = frmDest.ImgSkin(0).Height
frmDest.ImgSkin(2).Visible = True

'*** BORDA INFERIOR ****
Load frmDest.ImgSkin(3)
frmDest.ImgSkin(3).Height = 80
frmDest.ImgSkin(3).Width = frmDest.Width
gtiObj.LikroTmuna ArqBinImg, "SKBT", ArqBinImgTmp
frmDest.ImgSkin(3).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(3).Picture = LoadPicture(sPathBin & "\SKBT.BMP")
frmDest.ImgSkin(3).Left = 0
frmDest.ImgSkin(3).Top = frmDest.Height - 80
frmDest.ImgSkin(3).Visible = True

'*** FECHAR UP ****
If bClose Then
    Load frmDest.ImgSkin(5)
    frmDest.ImgSkin(5).Height = 225
    frmDest.ImgSkin(5).Width = 225
    gtiObj.LikroTmuna ArqBinImg, "SKCU", ArqBinImgTmp
    frmDest.ImgSkin(5).Picture = LoadPicture(ArqBinImgTmp)
'    frmDest.ImgSkin(5).Picture = LoadPicture(sPathBin & "\SKCU.BMP")
    frmDest.ImgSkin(5).Left = frmDest.Width - 350
    frmDest.ImgSkin(5).Top = 60
    frmDest.ImgSkin(5).Visible = True
    frmDest.ImgSkin(5).ZOrder 0
Else
    GoTo fim
End If

'*** RESTAURAR UP ****
If bMaximize Then
    Load frmDest.ImgSkin(6)
    frmDest.ImgSkin(6).Height = 225
    frmDest.ImgSkin(6).Width = 225
    gtiObj.LikroTmuna ArqBinImg, "SKPU", ArqBinImgTmp
    frmDest.ImgSkin(6).Picture = LoadPicture(ArqBinImgTmp)
'    frmDest.ImgSkin(6).Picture = LoadPicture(sPathBin & "\SKPU.BMP")
    frmDest.ImgSkin(6).Left = frmDest.Width - 600
    frmDest.ImgSkin(6).Top = 60
    frmDest.ImgSkin(6).Visible = True
    frmDest.ImgSkin(6).ZOrder 0
End If

'*** MINIMIZAR UP ****
If bMinimize Then
    Load frmDest.ImgSkin(4)
    frmDest.ImgSkin(4).Height = 225
    frmDest.ImgSkin(4).Width = 225
    gtiObj.LikroTmuna ArqBinImg, "SKMU", ArqBinImgTmp
    frmDest.ImgSkin(4).Picture = LoadPicture(ArqBinImgTmp)
'    frmDest.ImgSkin(4).Picture = LoadPicture(sPathBin & "\SKMU.BMP")
    If bMaximize Then
         frmDest.ImgSkin(4).Left = frmDest.Width - 850
    Else
        frmDest.ImgSkin(4).Left = frmDest.Width - 600
    End If
    frmDest.ImgSkin(4).Top = 60
    frmDest.ImgSkin(4).Visible = True
    frmDest.ImgSkin(4).ZOrder 0
End If

fim:
'*** ICONE ***
Load frmDest.ImgSkin(7)
frmDest.ImgSkin(7).Height = 225
frmDest.ImgSkin(7).Width = 225
frmDest.ImgSkin(7).Picture = frmDest.Icon
frmDest.ImgSkin(7).Left = 80
frmDest.ImgSkin(7).Top = 50
frmDest.ImgSkin(7).Visible = True
frmDest.ImgSkin(7).ZOrder 0

frmDest.lblSkin.Caption = frmDest.Caption
frmDest.lblSkin.Alignment = vbCenter
frmDest.lblSkin.BorderStyle = 0
frmDest.lblSkin.Move 60, 60, frmDest.Width - 700, 200
frmDest.lblSkin.ZOrder 0

End Sub

Public Sub CloseForm(frmDest As Form)
Unload frmDest
End Sub

Public Sub FormDown(frmDest As Form)
frmDest.WindowState = vbMinimized
ResizeForm frmDest
End Sub

Public Sub FormUP(frmDest As Form)
frmDest.WindowState = vbMaximized
ResizeForm frmDest
End Sub

Public Sub ShowMDown(frmDest As Form)
gtiObj.LikroTmuna ArqBinImg, "SKMD", ArqBinImgTmp
frmDest.ImgSkin(4).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(4).Picture = LoadPicture(sPathBin & "\SKMD.BMP")
'frmDest.ImgSkin(4).Picture = LoadPicture(sPathImage & "\MINUSDOWN.JPG")
End Sub

Public Sub ShowMUP(frmDest As Form)
gtiObj.LikroTmuna ArqBinImg, "SKMU", ArqBinImgTmp
frmDest.ImgSkin(4).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(4).Picture = LoadPicture(sPathBin & "\SKMU.BMP")
'frmDest.ImgSkin(4).Picture = LoadPicture(sPathImage & "\MINUSUP.JPG")
End Sub

Public Sub ShowCDown(frmDest As Form)
gtiObj.LikroTmuna ArqBinImg, "SKCD", ArqBinImgTmp
frmDest.ImgSkin(5).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(5).Picture = LoadPicture(sPathBin & "\SKCD.BMP")
'frmDest.ImgSkin(5).Picture = LoadPicture(sPathImage & "\CLOSEDOWN.JPG")
End Sub

Public Sub ShowCUP(frmDest As Form)
gtiObj.LikroTmuna ArqBinImg, "SKCU", ArqBinImgTmp
frmDest.ImgSkin(5).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(5).Picture = LoadPicture(sPathBin & "\SKCU.BMP")
'frmDest.ImgSkin(5).Picture = LoadPicture(sPathImage & "\CLOSEUP.JPG")
End Sub

Public Sub ShowRup(frmDest As Form)
gtiObj.LikroTmuna ArqBinImg, "SKPU", ArqBinImgTmp
frmDest.ImgSkin(6).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(6).Picture = LoadPicture(sPathBin & "\SKPU.BMP")
'frmDest.ImgSkin(6).Picture = LoadPicture(sPathImage & "\PLUSUP.JPG")
End Sub

Public Sub ShowRDown(frmDest As Form)
gtiObj.LikroTmuna ArqBinImg, "SKPD", ArqBinImgTmp
frmDest.ImgSkin(6).Picture = LoadPicture(ArqBinImgTmp)
'frmDest.ImgSkin(6).Picture = LoadPicture(sPathBin & "\SKPD.BMP")
'frmDest.ImgSkin(6).Picture = LoadPicture(sPathImage & "\PLUSDOWN.JPG")
End Sub

Public Sub ResizeForm(frmDest As Form)
On Error Resume Next
If bSkin Then
Dim bMinimize As Boolean, bMaximize As Boolean, bClose As Boolean
bMinimize = frmDest.MinButton
bMaximize = frmDest.MaxButton
bClose = frmDest.ControlBox

'*** BARRA DE TITULO ****
frmDest.ImgSkin(0).Move 0, 0, frmDest.Width, 330
'*** BORDA ESQUERDA ****
frmDest.ImgSkin(1).Move 0, frmDest.ImgSkin(0).Height, 60, frmDest.Height - frmDest.ImgSkin(0).Height
''*** BORDA DIREITA ****
frmDest.ImgSkin(2).Move frmDest.Width - 80, frmDest.ImgSkin(0).Height, 80, frmDest.Height - frmDest.ImgSkin(0).Height
''*** BORDA INFERIOR ****
frmDest.ImgSkin(3).Move 0, frmDest.Height - 80, frmDest.Width, 80
''*** FECHAR UP ****
If bClose Then
     frmDest.ImgSkin(5).Move frmDest.Width - 350, 60, 225, 225
Else
    GoTo Fim2
End If
'*** RESTAURAR UP ****
If bMaximize Then
    frmDest.ImgSkin(6).Move frmDest.Width - 600, 60, 225, 225
End If
''*** MINIMIZAR UP ****
If bMinimize Then
    If bMaximize Then
         frmDest.ImgSkin(4).Move frmDest.Width - 850, 60, 225, 225
         frmDest.ImgSkin(4).Left = frmDest.Width - 850
    Else
        frmDest.ImgSkin(4).Move frmDest.Width - 600, 60, 225, 225
        frmDest.ImgSkin(4).Left = frmDest.Width - 600
    End If
End If
frmDest.lblSkin.Move 60, 60, frmDest.Width - 700, 200
Refresh frmDest
Else
   frmDest.lblSkin.Visible = False
End If
Fim2:
End Sub


