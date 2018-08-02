VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmResumoImobiliario 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo de Imóveis Cadastrados"
   ClientHeight    =   3945
   ClientLeft      =   2265
   ClientTop       =   4500
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   9450
   Begin Tributacao.XP_ProgressBar Pb 
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   3555
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12500670
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   360
      Left            =   8040
      TabIndex        =   2
      ToolTipText     =   "Sair da Tela"
      Top             =   3480
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmResumoImobiliario.frx":0000
      PICN            =   "frmResumoImobiliario.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   6660
      TabIndex        =   1
      ToolTipText     =   "Cancelar Edição"
      Top             =   3480
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Executar"
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
      MICON           =   "frmResumoImobiliario.frx":008A
      PICN            =   "frmResumoImobiliario.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdResumo 
      Height          =   3315
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   5847
      _Version        =   393216
      Rows            =   11
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "frmResumoImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nContaPTAreaJB As Double
Dim nContaPUAreaJB As Double
Dim nContaOUAreaJB As Double
Dim nContaPTSemAreaJB As Double
Dim nContaPUSemAreaJB As Double
Dim nContaOUSemAreaJB As Double
Dim nSomaPTAreaJB As Double
Dim nSomaPUAreaJB As Double
Dim nSomaOUAreaJB As Double
Dim nContaPTAreaCR As Double
Dim nContaPUAreaCR As Double
Dim nContaOUAreaCR As Double
Dim nContaPTSemAreaCR As Double
Dim nContaPUSemAreaCR As Double
Dim nContaOUSemAreaCR As Double
Dim nSomaPTAreaCR As Double
Dim nSomaPUAreaCR As Double
Dim nSomaOUAreaCR As Double
Dim nContaPTAreaLZ As Double
Dim nContaPUAreaLZ As Double
Dim nContaOUAreaLZ As Double
Dim nContaPTSemAreaLZ As Double
Dim nContaPUSemAreaLZ As Double
Dim nContaOUSemAreaLZ As Double
Dim nSomaPTAreaLZ As Double
Dim nSomaPUAreaLZ As Double
Dim nSomaOUAreaLZ As Double
Dim nSomaTotalArea As Double
Dim nContaTotalArea As Double
Dim nContaTotalSemArea As Double
Dim nSomaSumario As Double
Dim nTotalJB As Double, nTotalCR As Double, nTotalLZ As Double

Private Sub cmdExec_Click()
Dim Tot As Double, nPos As Double, sCateg As String, nSomaArea As Double, nDist As Double, nOldCod As Double, nCodReduz As Double

Ocupado

nContaPTAreaJB = 0: nContaPUAreaJB = 0: nContaOUAreaJB = 0: nContaPTSemAreaJB = 0: nContaPUSemAreaJB = 0: nContaOUSemAreaJB = 0:
nSomaPTAreaJB = 0: nSomaPUAreaJB = 0: nSomaOUAreaJB = 0: nContaPTAreaCR = 0: nContaPUAreaCR = 0: nContaOUAreaCR = 0: nContaPTSemAreaCR = 0:
nContaPUSemAreaCR = 0: nContaOUSemAreaCR = 0: nSomaPTAreaCR = 0: nSomaPUAreaCR = 0: nSomaOUAreaCR = 0: nContaPTAreaLZ = 0: nContaPUAreaLZ = 0:
nContaOUAreaLZ = 0: nContaPTSemAreaLZ = 0: nContaPUSemAreaLZ = 0: nContaOUSemAreaLZ = 0: nSomaPTAreaLZ = 0: nSomaPUAreaLZ = 0: nSomaOUAreaLZ = 0:
nSomaTotalArea = 0: nContaTotalArea = 0: nContaTotalSemArea = 0

Pb.Value = 0: nOldCod = 0: nSomaArea = 0

Sql = "SELECT cadimob.codreduzido,SUM(areas.areaconstr) AS SomaArea, cadimob.distrito, cadimob.dt_codcategprop "
Sql = Sql & "FROM cadimob LEFT OUTER JOIN areas ON cadimob.codreduzido = areas.codreduzido Where cadimob.Inativo = 0 "
Sql = Sql & "GROUP BY cadimob.codreduzido, cadimob.distrito, cadimob.dt_codcategprop ORDER BY cadimob.codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        
        nCodCateg = !Dt_CodCategProp
        nDist = !Distrito
        Select Case nCodCateg
            Case 1
                sCateg = "A"
            Case 2, 3, 4
                sCateg = "U"
            Case Else
                sCateg = "O"
        End Select
        
        If Not IsNull(!SOMAAREA) Then
            nSomaArea = !SOMAAREA
        Else
            nSomaArea = 0
        End If
        nCodReduz = !CODREDUZIDO
        If nSomaArea > 0 Then
            nSomaTotalArea = nSomaTotalArea + nSomaArea
            nContaTotalArea = nContaTotalArea + 1
            If nDist = 1 Then
                If sCateg = "A" Then
                    nSomaPTAreaJB = nSomaPTAreaJB + nSomaArea
                    nContaPTAreaJB = nContaPTAreaJB + 1
                ElseIf sCateg = "U" Then
                    nSomaPUAreaJB = nSomaPUAreaJB + nSomaArea
                    nContaPUAreaJB = nContaPUAreaJB + 1
                Else
                    nSomaOUAreaJB = nSomaOUAreaJB + nSomaArea
                    nContaOUAreaJB = nContaOUAreaJB + 1
                End If
            ElseIf nDist = 2 Then
                If sCateg = "A" Then
                    nSomaPTAreaCR = nSomaPTAreaCR + nSomaArea
                    nContaPTAreaCR = nContaPTAreaCR + 1
                ElseIf sCateg = "U" Then
                    nSomaPUAreaCR = nSomaPUAreaCR + nSomaArea
                    nContaPUAreaCR = nContaPUAreaCR + 1
                Else
                    nSomaOUAreaCR = nSomaOUAreaCR + nSomaArea
                    nContaOUAreaCR = nContaOUAreaCR + 1
                End If
            ElseIf nDist = 3 Then
                If sCateg = "A" Then
                    nSomaPTAreaLZ = nSomaPTAreaLZ + nSomaArea
                    nContaPTAreaLZ = nContaPTAreaLZ + 1
                ElseIf sCateg = "U" Then
                    nSomaPUAreaLZ = nSomaPUAreaLZ + nSomaArea
                    nContaPUAreaLZ = nContaPUAreaLZ + 1
                Else
                    nSomaOUAreaLZ = nSomaOUAreaLZ + nSomaArea
                    nContaOUAreaLZ = nContaOUAreaLZ + 1
                End If
            End If
        Else
            nContaTotalSemArea = nContaTotalSemArea + 1
            If nDist = 1 Then
                If sCateg = "A" Then
                    nContaPTSemAreaJB = nContaPTSemAreaJB + 1
                ElseIf sCateg = "U" Then
                    nContaPUSemAreaJB = nContaPUSemAreaJB + 1
                Else
                    nContaOUSemAreaJB = nContaOUSemAreaJB + 1
                End If
            ElseIf nDist = 2 Then
                If sCateg = "A" Then
                    nContaPTSemAreaCR = nContaPTSemAreaCR + 1
                ElseIf sCateg = "U" Then
                    nContaPUSemAreaCR = nContaPUSemAreaCR + 1
                Else
                    nContaOUSemAreaCR = nContaOUSemAreaCR + 1
                End If
            ElseIf nDist = 3 Then
                If sCateg = "A" Then
                    nContaPTSemAreaLZ = nContaPTSemAreaLZ + 1
                ElseIf sCateg = "U" Then
                    nContaPUSemAreaLZ = nContaPUSemAreaLZ + 1
                Else
                    nContaOUSemAreaLZ = nContaOUSemAreaLZ + 1
                End If
            End If
        End If
       
        nPos = .AbsolutePosition
        If nPos Mod 100 = 0 Then
            CallPb nPos, CLng(nTot)
            FillTotal
        End If
       
       .MoveNext
    Loop
   .Close
End With
Liberado
Pb.Value = 100

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
PreparaGrid
End Sub

Private Sub FillTotal()


With grdResumo
    .TextMatrix(1, 2) = FormatNumber(nSomaPTAreaJB, 2)
    .TextMatrix(2, 2) = FormatNumber(nSomaPUAreaJB, 2)
    .TextMatrix(3, 2) = FormatNumber(nSomaOUAreaJB, 2)
    .TextMatrix(4, 2) = FormatNumber(nSomaPUAreaCR, 2)
    .TextMatrix(5, 2) = FormatNumber(nSomaPUAreaCR, 2)
    .TextMatrix(6, 2) = FormatNumber(nSomaOUAreaCR, 2)
    .TextMatrix(7, 2) = FormatNumber(nSomaPTAreaLZ, 2)
    .TextMatrix(8, 2) = FormatNumber(nSomaPUAreaLZ, 2)
    .TextMatrix(9, 2) = FormatNumber(nSomaOUAreaLZ, 2)
    .TextMatrix(10, 2) = FormatNumber(nSomaTotalArea, 2)
    .TextMatrix(1, 3) = nContaPTAreaJB
    .TextMatrix(2, 3) = nContaPUAreaJB
    .TextMatrix(3, 3) = nContaOUAreaJB
    .TextMatrix(4, 3) = nContaPTAreaCR
    .TextMatrix(5, 3) = nContaPUAreaCR
    .TextMatrix(6, 3) = nContaOUAreaCR
    .TextMatrix(7, 3) = nContaPTAreaLZ
    .TextMatrix(8, 3) = nContaPUAreaLZ
    .TextMatrix(9, 3) = nContaOUAreaLZ
    .TextMatrix(10, 3) = nContaTotalArea
    .TextMatrix(1, 4) = nContaPTSemAreaJB
    .TextMatrix(2, 4) = nContaPUSemAreaJB
    .TextMatrix(3, 4) = nContaOUSemAreaJB
    .TextMatrix(4, 4) = nContaPTSemAreaCR
    .TextMatrix(5, 4) = nContaPUSemAreaCR
    .TextMatrix(6, 4) = nContaOUSemAreaCR
    .TextMatrix(7, 4) = nContaPTSemAreaLZ
    .TextMatrix(8, 4) = nContaPUSemAreaLZ
    .TextMatrix(9, 4) = nContaOUSemAreaLZ
    .TextMatrix(10, 4) = nContaTotalSemArea
    'sumario
    .TextMatrix(1, 5) = nContaPTAreaJB + nContaPTSemAreaJB
    .TextMatrix(2, 5) = nContaPUAreaJB + nContaPUSemAreaJB
    .TextMatrix(3, 5) = nContaOUAreaJB + nContaOUSemAreaJB
    .TextMatrix(4, 5) = nContaPTAreaCR + nContaPTSemAreaCR
    .TextMatrix(5, 5) = nContaPUAreaCR + nContaPUSemAreaCR
    .TextMatrix(6, 5) = nContaOUAreaCR + nContaOUSemAreaCR
    .TextMatrix(7, 5) = nContaPTAreaLZ + nContaPTSemAreaLZ
    .TextMatrix(8, 5) = nContaPUAreaLZ + nContaPUSemAreaLZ
    .TextMatrix(9, 5) = nContaOUAreaLZ + nContaOUSemAreaLZ
    
    nSomaSumario = 0
    For x = 1 To 9
        nSomaSumario = nSomaSumario + (.TextMatrix(x, 5))
    Next
    .TextMatrix(10, 5) = nSomaSumario
    
    nTotalJB = nContaPTAreaJB + nContaPTSemAreaJB + nContaPUAreaJB + nContaPUSemAreaJB + nContaOUAreaJB + nContaOUSemAreaJB
    nTotalCR = nContaPTAreaCR + nContaPTSemAreaCR + nContaPUAreaCR + nContaPUSemAreaCR + nContaOUAreaCR + nContaOUSemAreaCR
    nTotalLZ = nContaPTAreaLZ + nContaPTSemAreaLZ + nContaPUAreaLZ + nContaPUSemAreaLZ + nContaOUAreaLZ + nContaOUSemAreaLZ
    .TextMatrix(1, 6) = "JB-> " & nTotalJB
    .TextMatrix(2, 6) = "JB-> " & nTotalJB
    .TextMatrix(3, 6) = "JB-> " & nTotalJB
    .TextMatrix(4, 6) = "CR-> " & nTotalCR
    .TextMatrix(5, 6) = "CR-> " & nTotalCR
    .TextMatrix(6, 6) = "CR-> " & nTotalCR
    .TextMatrix(7, 6) = "LZ-> " & nTotalLZ
    .TextMatrix(8, 6) = "LZ-> " & nTotalLZ
    .TextMatrix(9, 6) = "LZ-> " & nTotalLZ
    
End With

End Sub

Private Sub PreparaGrid()
Dim x As Double, y As Double
With grdResumo
    .TextMatrix(0, 0) = "DISTRITO"
    .TextMatrix(1, 0) = "JABOTICABAL"
    .TextMatrix(2, 0) = "JABOTICABAL"
    .TextMatrix(3, 0) = "JABOTICABAL"
    .TextMatrix(4, 0) = "CÓRR. RICO"
    .TextMatrix(5, 0) = "CÓRR. RICO"
    .TextMatrix(6, 0) = "CÓRR. RICO"
    .TextMatrix(7, 0) = "LUZITÂNIA"
    .TextMatrix(8, 0) = "LUZITÂNIA"
    .TextMatrix(9, 0) = "LUZITÂNIA"
    .TextMatrix(10, 0) = "TOTAL"
    .TextMatrix(0, 1) = "CATEGORIA"
    .TextMatrix(1, 1) = "PARTICULAR"
    .TextMatrix(2, 1) = "PÚBLICO"
    .TextMatrix(3, 1) = "OUTROS"
    .TextMatrix(4, 1) = "PARTICULAR"
    .TextMatrix(5, 1) = "PÚBLICO"
    .TextMatrix(6, 1) = "OUTROS"
    .TextMatrix(7, 1) = "PARTICULAR"
    .TextMatrix(8, 1) = "PÚBLICO"
    .TextMatrix(9, 1) = "OUTROS"
    .TextMatrix(0, 2) = "SOMA ÁREA"
    .TextMatrix(0, 3) = "COM ÁREA"
    .TextMatrix(0, 4) = "SEM ÁREA"
    .TextMatrix(0, 5) = "SUMÁRIO"
    .TextMatrix(0, 6) = "TOTAL"

    For y = 2 To 6
        For x = 1 To 10
            .col = y: .Row = x
            If y = 2 Then
                .Text = "0,00"
            Else
                .Text = "0"
            End If
            If x = 10 Then
                .CellFontBold = True
            End If
        Next
    Next
    .Row = 0
    For x = 0 To 6
        .col = x
        .CellFontBold = True
        .CellBackColor = vbRed
        .CellForeColor = Branco
    Next
    .col = 0
    For x = 1 To 9
        .Row = x
        .CellFontBold = True
        .CellBackColor = &H800000
        .CellForeColor = Branco
    Next
    For y = 1 To 5
        .col = y
        For x = 1 To 9 Step 2
            .Row = x
            .CellBackColor = &H95D1FB
        Next
    Next
    For y = 1 To 5
        .col = y
        For x = 2 To 9 Step 2
            .Row = x
            .CellBackColor = &H2FF7F3
        Next
    Next
    .col = 6
    For x = 1 To 9
        .Row = x
        .CellFontBold = True
        .CellBackColor = &H81FF80
        .CellForeColor = Preto
        If x < 4 Then
            .Text = "JB-> 0"
        ElseIf x > 3 And x < 6 Then
            .Text = "CR-> 0"
        ElseIf x > 5 Then
            .Text = "LZ-> 0"
        End If
    Next
    .Row = .Rows - 1
    For x = 0 To 6
        .col = x
        .CellFontBold = True
        .CellBackColor = &H81FF80
        .CellForeColor = Preto
        If x = 6 Then .Text = ""
    Next
    
    .ColWidth(0) = 1350: .ColWidth(1) = 1300
    .ColWidth(2) = 1300: .ColWidth(3) = 1300
    .ColWidth(4) = 1300: .ColWidth(5) = 1300: .ColWidth(6) = 1500
    .RowHeightMin = 300
    .MergeCells = flexMergeRestrictColumns
    .MergeCol(0) = True: .MergeCol(6) = True
    .ColAlignment(2) = flexAlignRightCenter: .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter: .ColAlignment(5) = flexAlignRightCenter: .ColAlignment(6) = flexAlignRightCenter
End With


End Sub

Private Sub CallPb(nPosF As Double, nTotal As Double)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.Value = (nPosF * 100) / nTotal
Else
   Pb.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub


