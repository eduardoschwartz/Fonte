VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSenhaStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Senhas"
   ClientHeight    =   5910
   ClientLeft      =   4065
   ClientTop       =   3300
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   8310
   Begin Tributacao.jcFrames frTela 
      Height          =   4740
      Index           =   0
      Left            =   45
      Top             =   630
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8361
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Senhas chamadas por guiche e banda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSFlexGridLib.MSFlexGrid grdMain 
         Height          =   4245
         Left            =   45
         TabIndex        =   8
         Top             =   450
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   7488
         _Version        =   393216
         Rows            =   14
         Cols            =   8
         RowHeightMin    =   300
         BackColorFixed  =   4194304
         ForeColorFixed  =   12648447
         BackColorSel    =   16777215
         ForeColorSel    =   8388608
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         GridLinesFixed  =   0
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "                 |^Pref.      |^2ª Via    |^Prefer.    |^Pat        |^REFIS   |^Pr.Refis |^Total         "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin prjChameleon.chameleonButton cmdRefresh 
      Height          =   315
      Left            =   2610
      TabIndex        =   12
      ToolTipText     =   "Atualizar lista"
      Top             =   5490
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "At&ualizar"
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
      MICON           =   "frmSenhaStatus.frx":0000
      PICN            =   "frmSenhaStatus.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtData 
      Height          =   330
      Left            =   765
      TabIndex        =   2
      Top             =   135
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      _Version        =   393216
      Format          =   161677313
      CurrentDate     =   40414
   End
   Begin prjChameleon.chameleonButton cmdAnterior 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Tela Anterior"
      Top             =   5490
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Anterior"
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
      MICON           =   "frmSenhaStatus.frx":012E
      PICN            =   "frmSenhaStatus.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdProximo 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      ToolTipText     =   "Próxima Tela"
      Top             =   5490
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Próximo"
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
      MICON           =   "frmSenhaStatus.frx":02A4
      PICN            =   "frmSenhaStatus.frx":02C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.jcFrames frTela 
      Height          =   4740
      Index           =   1
      Left            =   60
      Top             =   630
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8361
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Senhas emitidas por hora"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSFlexGridLib.MSFlexGrid grdHora 
         Height          =   3930
         Left            =   45
         TabIndex        =   9
         Top             =   450
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   6932
         _Version        =   393216
         Rows            =   13
         Cols            =   4
         RowHeightMin    =   300
         BackColorFixed  =   4194304
         ForeColorFixed  =   12648447
         BackColorSel    =   16777215
         ForeColorSel    =   8388608
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         GridLinesFixed  =   0
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "^                      |^Emitidas     |^Segundos|^Espera Méd"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEsperaMedia 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2880
         TabIndex        =   11
         Top             =   4410
         Width           =   1860
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo médio de espera: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   4410
         Width           =   2760
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Na fila de espera..:"
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   7
      Top             =   180
      Width           =   1410
   End
   Begin VB.Label lblSenhaEspera 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   6075
      TabIndex        =   6
      Top             =   180
      Width           =   645
   End
   Begin VB.Label lblSenhaGerada 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3870
      TabIndex        =   5
      Top             =   180
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Senhas Geradas..:"
      Height          =   195
      Index           =   1
      Left            =   2475
      TabIndex        =   4
      Top             =   180
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data...:"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   180
      Width           =   645
   End
End
Attribute VB_Name = "frmSenhaStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SENHA
    sHoraEntrada As String
    sHoraChamada As String
    nBanda As Integer
    nGuiche As Integer
    nSenha As Integer
End Type

Dim nTela As Integer

Private Sub cmdAnterior_Click()
nTela = nTela - 1
MudaTela
End Sub

Private Sub cmdProximo_Click()
nTela = nTela + 1
MudaTela
End Sub

Private Sub cmdRefresh_Click()
Atualiza
End Sub

Private Sub dtData_Change()
Atualiza
End Sub

Private Sub Form_Load()
Dim x As Integer
Centraliza Me
nTela = 0
With grdMain
    For x = 1 To 12
        .TextMatrix(x, 0) = "Guiche " & Format(x, "00")
    Next
    .TextMatrix(.Rows - 1, 0) = "Total"
    .Row = .Rows - 1
    .col = 0
    .CellBackColor = vbRed
    .Row = 0
    .col = .Cols - 1
    .CellBackColor = vbRed
End With

With grdHora
    .COLWIDTH(2) = 0
    .Row = 0
    .col = 2
    .CellForeColor = vbBlue
    For x = 1 To 12
        .TextMatrix(x, 0) = Format(6 + x, "00") & ":00 às " & Format(7 + x, "00") & ":00"
    Next
End With

MudaTela
dtData.value = Now
Atualiza
End Sub

Private Sub Atualiza()
Dim Sql As String, RdoAux As rdoResultset, aSenha() As SENHA, nPos As Integer, nTotal3 As Integer
Dim a As Integer, b As Integer, aTotal(12, 6) As Integer, nTotal As Long, nTotal2 As Long
Dim aHora(12, 4) As Long, nDif As Integer

ReDim aSenha(0)

Me.Caption = "Controle de Senhas às " & Format(Now, "hh:mm:ss")
nTotal3 = 0

Sql = "SELECT * FROM SSPAC WHERE DATAENTRADA='" & Format(dtData.value, "mm/dd/yyyy") & "' ORDER BY SENHA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nPos = UBound(aSenha) + 1
        ReDim Preserve aSenha(nPos)
        aSenha(nPos).sHoraEntrada = !HORAENTRADA
        aSenha(nPos).sHoraChamada = SubNull(!horachamada)
        aSenha(nPos).nSenha = !SENHA
        aSenha(nPos).nBanda = !BANDA
        aSenha(nPos).nGuiche = Val(SubNull(!GUICHE))
        nTotal3 = nTotal3 + 1
        If Not IsNull(!GUICHE) Then
            If !GUICHE > 0 And !GUICHE <= 12 Then
                aTotal(!GUICHE, !BANDA) = aTotal(!GUICHE, !BANDA) + 1
            End If
        End If
       .MoveNext
    Loop
   .Close
End With

lblSenhaGerada.Caption = nTotal3

For a = 1 To 12
    For b = 1 To 6
        grdMain.TextMatrix(a, b) = ""
    Next
Next

With grdMain
    For a = 1 To 12
        nTotal = 0
        For b = 1 To 6
           .TextMatrix(a, b) = aTotal(a, b)
            nTotal = nTotal + aTotal(a, b)
        Next
        .col = .Cols - 1
        .Row = a
        .CellForeColor = vbRed
        .Text = nTotal
    Next
    nTotal2 = 0
    For a = 1 To 6
        nTotal = 0
        For b = 1 To 12
            nTotal = nTotal + Val(.TextMatrix(b, a))
        Next
        .col = a
        .Row = b
        .CellForeColor = vbRed
        .Text = nTotal
        nTotal2 = nTotal2 + nTotal
    Next
    .col = .Cols - 1
    .Row = .Rows - 1
    .CellForeColor = vbRed
    .Text = nTotal2
    .col = 1
    .Row = 1
End With

lblSenhaEspera.Caption = nTotal3 - nTotal2

For nPos = 1 To UBound(aSenha)
    With aSenha(nPos)
        If CDate(.sHoraEntrada) >= CDate("07:00:00") And CDate(.sHoraEntrada) < CDate("08:00:00") Then
            aHora(1, 1) = aHora(1, 1) + 1
            If .sHoraChamada = "" Then
                aHora(1, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(1, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(1, 3) = aHora(1, 3) + aHora(1, 2)
            End If
            aHora(1, 4) = aHora(1, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("08:00:00") And CDate(.sHoraEntrada) < CDate("09:00:00") Then
            aHora(2, 1) = aHora(2, 1) + 1
            If .sHoraChamada = "" Then
                aHora(2, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(2, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(2, 3) = aHora(2, 3) + aHora(2, 2)
            End If
            aHora(2, 4) = aHora(2, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("09:00:00") And CDate(.sHoraEntrada) < CDate("10:00:00") Then
            aHora(3, 1) = aHora(3, 1) + 1
            If .sHoraChamada = "" Then
                aHora(3, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(3, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(3, 3) = aHora(3, 3) + aHora(3, 2)
            End If
            aHora(3, 4) = aHora(3, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("10:00:00") And CDate(.sHoraEntrada) < CDate("11:00:00") Then
            aHora(4, 1) = aHora(4, 1) + 1
            If .sHoraChamada = "" Then
                aHora(4, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(4, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(4, 3) = aHora(4, 3) + aHora(4, 2)
            End If
            aHora(4, 4) = aHora(4, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("11:00:00") And CDate(.sHoraEntrada) < CDate("12:00:00") Then
            aHora(5, 1) = aHora(5, 1) + 1
            If .sHoraChamada = "" Then
                aHora(5, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(5, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(5, 3) = aHora(5, 3) + aHora(5, 2)
            End If
            aHora(5, 4) = aHora(5, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("12:00:00") And CDate(.sHoraEntrada) < CDate("13:00:00") Then
            aHora(6, 1) = aHora(6, 1) + 1
            If .sHoraChamada = "" Then
                aHora(6, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(6, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(6, 3) = aHora(6, 3) + aHora(6, 2)
            End If
            aHora(6, 4) = aHora(6, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("13:00:00") And CDate(.sHoraEntrada) < CDate("14:00:00") Then
            aHora(7, 1) = aHora(7, 1) + 1
            If .sHoraChamada = "" Then
                aHora(7, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(7, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(7, 3) = aHora(7, 3) + aHora(7, 2)
            End If
            aHora(7, 4) = aHora(7, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("14:00:00") And CDate(.sHoraEntrada) < CDate("15:00:00") Then
            aHora(8, 1) = aHora(8, 1) + 1
            If .sHoraChamada = "" Then
                aHora(8, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(8, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(8, 3) = aHora(8, 3) + aHora(8, 2)
            End If
            aHora(8, 4) = aHora(8, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("15:00:00") And CDate(.sHoraEntrada) < CDate("16:00:00") Then
            aHora(9, 1) = aHora(9, 1) + 1
            If .sHoraChamada = "" Then
                aHora(9, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(9, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(9, 3) = aHora(9, 3) + aHora(9, 2)
            End If
            aHora(9, 4) = aHora(9, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("16:00:00") And CDate(.sHoraEntrada) < CDate("17:00:00") Then
            aHora(10, 1) = aHora(10, 1) + 1
            If .sHoraChamada = "" Then
                aHora(10, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(10, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(10, 3) = aHora(10, 3) + aHora(10, 2)
            End If
            aHora(10, 4) = aHora(10, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("17:00:00") And CDate(.sHoraEntrada) < CDate("18:00:00") Then
            aHora(11, 1) = aHora(11, 1) + 1
            If .sHoraChamada = "" Then
                aHora(11, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(11, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(11, 3) = aHora(11, 3) + aHora(11, 2)
            End If
            aHora(11, 4) = aHora(11, 4) + 1
        ElseIf CDate(.sHoraEntrada) >= CDate("18:00:00") And CDate(.sHoraEntrada) < CDate("19:00:00") Then
            aHora(12, 1) = aHora(12, 1) + 1
            If .sHoraChamada = "" Then
                aHora(12, 2) = Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(Format(Now, "hh:mm:ss"))))
            Else
                aHora(12, 2) = IIf(.sHoraChamada = "", 0, Abs(DateDiff("s", CDate(.sHoraEntrada), CDate(.sHoraChamada))))
                aHora(12, 3) = aHora(12, 3) + aHora(12, 2)
            End If
            aHora(12, 4) = aHora(12, 4) + 1
        End If
    End With
Next

For a = 1 To 12
    For b = 1 To 2
        grdHora.TextMatrix(a, b) = ""
    Next
Next

With grdHora
    For a = 1 To 12
        .Row = a
        .col = 2
        .CellForeColor = vbWhite
        If aHora(a, 3) > 0 Then
            .TextMatrix(a, 3) = SecondToTime(aHora(a, 3) / aHora(a, 4))
        Else
            .TextMatrix(a, 3) = 0
        End If
    Next
End With

nTotal = 0: nTotal2 = 0
With grdHora
    For a = 1 To 12
        .Row = a
        .col = 2
        .CellForeColor = vbWhite
        .TextMatrix(a, 1) = aHora(a, 4)
        nTotal = nTotal + aHora(a, 3)
        nTotal2 = nTotal2 + aHora(a, 4)
    Next
    .col = 1
    .Row = 1
End With

If nTotal2 > 0 Then
    lblEsperaMedia.Caption = SecondToTime(nTotal / nTotal2)
Else
    lblEsperaMedia.Caption = "00:00:00"
End If

End Sub

Private Sub MudaTela()
Dim c As Integer

For c = 0 To 1
    If c = nTela Then
        frTela(c).Visible = True
    Else
        frTela(c).Visible = False
    End If
Next

If nTela = 0 Then
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = True
ElseIf nTela = 1 Then
    cmdAnterior.Enabled = True
    cmdProximo.Enabled = False
Else
    cmdAnterior.Enabled = True
    cmdProximo.Enabled = True
End If

End Sub

Private Function SecondToTime(nSeconds As Integer) As String
Dim Daycount As Long, SecondsRemaining As Integer, HourCount As Integer
Dim MinutesCount As Integer, SecondsCount As Integer

Daycount = nSeconds \ (86400)
SecondsRemaining = nSeconds - (Daycount * (86400))
HourCount = Abs(SecondsRemaining \ (60 * 60))
SecondsRemaining = Abs(SecondsRemaining Mod (60 * 60))
MinutesCount = Abs(SecondsRemaining \ 60)
SecondsCount = Abs(SecondsRemaining Mod 60)
SecondToTime = Format(HourCount, "00") & ":" & Format(MinutesCount, "00") & ":" & Format(SecondsCount, "00")
End Function
