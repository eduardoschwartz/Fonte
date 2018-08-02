VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGeraDebito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração Manual de Débitos"
   ClientHeight    =   3900
   ClientLeft      =   3240
   ClientTop       =   3150
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   8595
   Begin VB.TextBox txtBase 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7500
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin prjChameleon.chameleonButton cmdAddTrib 
      Height          =   315
      Left            =   6390
      TabIndex        =   29
      ToolTipText     =   "Novo Registro"
      Top             =   3150
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Adicionar"
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
      MICON           =   "frmGeraDebito.frx":0000
      PICN            =   "frmGeraDebito.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin prjChameleon.chameleonButton cmdDelParc 
      Height          =   315
      Left            =   7470
      TabIndex        =   12
      ToolTipText     =   "Excluir Registro"
      Top             =   1860
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmGeraDebito.frx":0176
      PICN            =   "frmGeraDebito.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtParcAte 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8100
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtParcDe 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7500
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtCompl 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7500
      TabIndex        =   6
      Top             =   450
      Width           =   975
   End
   Begin VB.TextBox txtVencto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7500
      TabIndex        =   8
      Top             =   1110
      Width           =   975
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7500
      TabIndex        =   7
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txtSeq 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      TabIndex        =   3
      Top             =   1110
      Width           =   975
   End
   Begin VB.TextBox txtLanc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      TabIndex        =   2
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      TabIndex        =   1
      Top             =   450
      Width           =   975
   End
   Begin prjChameleon.chameleonButton cmdAddParc 
      Height          =   315
      Left            =   6390
      TabIndex        =   11
      ToolTipText     =   "Novo Registro"
      Top             =   1845
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Adicionar"
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
      MICON           =   "frmGeraDebito.frx":0234
      PICN            =   "frmGeraDebito.frx":0250
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Height          =   315
      Left            =   5310
      TabIndex        =   10
      ToolTipText     =   "Gerar Parcelas"
      Top             =   1860
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Gerar"
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
      MICON           =   "frmGeraDebito.frx":03AA
      PICN            =   "frmGeraDebito.frx":03C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDelTrib 
      Height          =   315
      Left            =   7470
      TabIndex        =   15
      ToolTipText     =   "Excluir Registro"
      Top             =   3150
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmGeraDebito.frx":0442
      PICN            =   "frmGeraDebito.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbTrib 
      Height          =   315
      Left            =   4950
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2730
      Width           =   3555
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4950
      TabIndex        =   14
      Top             =   3120
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdParcela 
      Height          =   2235
      Left            =   30
      TabIndex        =   16
      Top             =   30
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   3942
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Ano     |^Lanc |^Seq|^Parc |^Compl |^Sit   |^Vencimento    "
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   1515
      Left            =   30
      TabIndex        =   17
      Top             =   2310
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   2672
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "<Tributo                                  |>Valor            "
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Base"
      Height          =   255
      Left            =   6390
      TabIndex        =   30
      Top             =   1500
      Width           =   1125
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "à"
      Height          =   255
      Left            =   7830
      TabIndex        =   28
      Top             =   180
      Width           =   315
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
      Height          =   255
      Left            =   6390
      TabIndex        =   27
      Top             =   1170
      Width           =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcela"
      Height          =   255
      Left            =   6390
      TabIndex        =   26
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   6390
      TabIndex        =   25
      Top             =   855
      Width           =   1125
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
      Height          =   255
      Left            =   6390
      TabIndex        =   24
      Top             =   495
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sequencia"
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   1170
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lancamento"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   825
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Exercício"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   495
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   180
      Width           =   1005
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tributo:"
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   2820
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   3195
      Width           =   615
   End
End
Attribute VB_Name = "frmGeraDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset

Private Sub cmdAddParc_Click()
Dim x As Integer, Y As Integer, bAchou As Boolean
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nSeqNovo As Integer
Dim nParc As Integer, nCodReduz As Long, dVencto As Date

If Val(txtCod.Text) = 0 Or Val(txtAno.Text) = 0 Or Val(txtLanc.Text) = 0 Or txtParcDe.Text = "" Or txtParcAte.Text = "" Or Val(txtStatus) = 0 Then
    MsgBox "Preencha todos os campos.", vbCritical, "Atenção"
Else
    If Val(txtSeq.Text) = 0 Then txtSeq.Text = 0
    If Val(txtCompl.Text) = 0 Then txtCompl.Text = 0
    If Not IsDate(txtVencto.Text) Then
        MsgBox "Vencimento inválido.", vbCritical, "Atenção"
    Else
        dVencto = txtVencto.Text
        nCodReduz = Val(txtCod.Text)
        nAno = Val(txtAno.Text)
        nLanc = Val(txtLanc.Text)
        nCompl = Val(txtCompl.Text)
        If Val(txtParcDe.Text) > Val(txtParcAte.Text) Then
            MsgBox "Parcela Final menor que inicial.", vbCritical, "Atenção"
        Else
            For Y = Val(txtParcDe.Text) To Val(txtParcAte.Text)
                nSeqNovo = Val(txtSeq.Text)
                bAchou = False
                With grdParcela
                    For x = 1 To .Rows - 1
                        nAno = .TextMatrix(x, 0)
                        nLanc = .TextMatrix(x, 1)
                        nSeq = .TextMatrix(x, 2)
                        nParc = .TextMatrix(x, 3)
                        nCompl = .TextMatrix(x, 4)
                        If nAno = Val(txtAno.Text) And nLanc = Val(txtLanc.Text) And nSeq = nSeqNovo And nParc = Y And nCompl = Val(txtCompl.Text) Then
                            bAchou = True
                            MsgBox "Parcela " & Y & " já adicionada ao grid.", vbCritical, "Atenção"
                            Exit For
                        End If
                    Next
                    
                    
                    If Not bAchou Then
                        .AddItem Val(txtAno.Text) & Chr(9) & Val(txtLanc.Text) & Chr(9) & nSeqNovo & Chr(9) & Y & Chr(9) & nCompl & Chr(9) & txtStatus.Text & Chr(9) & dVencto
                         dVencto = Format(DateAdd("m", 1, dVencto))
                    End If
                End With
            Next
        End If
    End If
End If

End Sub

Private Sub cmdAddTrib_Click()
Dim x As Integer, bAchou As Boolean
If Trim(txtValor.Text) = "" Then txtValor.Text = 0

bAchou = False
For x = 1 To grdTrib.Rows - 1
    If grdTrib.TextMatrix(grdTrib.Row, 0) = cmbTrib.Text Then
        bAchou = True
    End If
Next
If bAchou Then
    MsgBox "Tributo já cadastrado.", vbExclamation, "Atenção"
    Exit Sub
Else
    grdTrib.AddItem cmbTrib.Text & Chr(9) & FormatNumber(txtValor.Text, 2)
End If

End Sub

Private Sub cmdDelParc_Click()
grdParcela.Rows = 1
End Sub

Private Sub cmdDelTrib_Click()
If grdTrib.Rows > 1 Then
    If grdTrib.Row > 0 Then
        If grdTrib.Rows > 2 Then
            grdTrib.RemoveItem (grdTrib.Row)
        Else
            grdTrib.Rows = 1
        End If
    End If
End If

End Sub

Private Sub cmdGerar_Click()
Dim x As Integer, Y As Integer

If grdTrib.Rows = 1 Then
    MsgBox "Selecione os tributos.", vbExclamation, "Atenção"
Else
    If MsgBox("Gerar estes débitos ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        For x = 1 To grdParcela.Rows - 1
            'GRAVA NA TABELA DEBITOPARCELA
'            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'            Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(txtLanc.Text) & "," & Val(txtSeq.Text) & "," & grdParcela.TextMatrix(x, 3) & "," & Val(txtCompl.Text) & "," & Val(txtStatus.Text) & ",'"
'            Sql = Sql & Format(grdParcela.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(txtBase.Text, "mm/dd/yyyy") & "',0" & ",'" & Left$(NomeDeLogin, 25) & "')"
            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
            Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(txtLanc.Text) & "," & Val(txtSeq.Text) & "," & grdParcela.TextMatrix(x, 3) & "," & Val(txtCompl.Text) & "," & Val(txtStatus.Text) & ",'"
            Sql = Sql & Format(grdParcela.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(txtBase.Text, "mm/dd/yyyy") & "',0" & "," & RetornaUsuarioID(NomeDeLogin) & ")"
            cn.Execute Sql, rdExecDirect
            For Y = 1 To grdTrib.Rows - 1
                'GRAVA NA TABELA DEBITO TRIBUTO
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(txtCod.Text) & "," & Val(txtAno.Text) & "," & Val(txtLanc.Text) & "," & Val(txtSeq.Text) & "," & grdParcela.TextMatrix(x, 3) & "," & Val(txtCompl.Text) & ","
                Sql = Sql & Val(Left$(grdTrib.TextMatrix(Y, 0), 3)) & "," & Virg2Ponto(RemovePonto(grdTrib.TextMatrix(Y, 1))) & ")"
                cn.Execute Sql, rdExecDirect
            Next
        Next
    End If
End If
MsgBox "Debitos Gerados", vbInformation, "Atenção"
grdParcela.Rows = 1

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaTributo
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub CarregaTributo()

Sql = "SELECT CODTRIBUTO,ABREVTRIBUTO FROM TRIBUTO ORDER BY ABREVTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTrib.AddItem Format(!CodTributo, "000") & "-" & !ABREVTRIBUTO
       .MoveNext
    Loop
   .Close
End With

End Sub

