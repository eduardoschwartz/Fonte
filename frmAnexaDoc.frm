VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAnexaDoc 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexação de Documento aos Lançamentos Selecionados "
   ClientHeight    =   5325
   ClientLeft      =   2460
   ClientTop       =   2145
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5640
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      Height          =   1050
      Left            =   30
      MaxLength       =   400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3750
      Width           =   5565
   End
   Begin prjChameleon.chameleonButton cmdHelp 
      Height          =   345
      Left            =   3315
      TabIndex        =   0
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   4905
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Ajuda"
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
      MICON           =   "frmAnexaDoc.frx":0000
      PICN            =   "frmAnexaDoc.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   4470
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   4905
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmAnexaDoc.frx":0176
      PICN            =   "frmAnexaDoc.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBaixa 
      Height          =   345
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Anexar nº de documento"
      Top             =   4905
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Anexar"
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
      MICON           =   "frmAnexaDoc.frx":0200
      PICN            =   "frmAnexaDoc.frx":021C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   2700
      Left            =   30
      TabIndex        =   3
      Top             =   45
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   1
      Cols            =   10
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
      FormatString    =   "^Ano      |^Lanc |^Seq|^Par  |^Com|^Sit   |^Vencimento  |^D |^A |>Principal     "
   End
   Begin prjChameleon.chameleonButton cmdCopy 
      Height          =   345
      Left            =   2685
      TabIndex        =   10
      ToolTipText     =   "Copiar nº de documento para a área de transferência"
      Top             =   2880
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmAnexaDoc.frx":0376
      PICN            =   "frmAnexaDoc.frx":0392
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblDataDoc 
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   4200
      TabIndex        =   9
      Top             =   2940
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Doc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   3210
      TabIndex        =   8
      Top             =   2940
      Width           =   960
   End
   Begin VB.Label lblNumDoc 
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   1155
      TabIndex        =   7
      Top             =   2940
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   75
      TabIndex        =   6
      Top             =   2940
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo da Anexação...:"
      Height          =   225
      Index           =   1
      Left            =   45
      TabIndex        =   5
      Top             =   3510
      Width           =   1935
   End
End
Attribute VB_Name = "frmAnexaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String
Dim sAno As String, sLanc As String, sSeq As String, sParc As String
Dim sComp As String, nCodReduz As Long

Private Sub cmdBaixa_Click()
Dim nLastDoc As Long, x As Integer

If Trim$(txtMotivo.Text) = "" Then
    MsgBox "Digite o Motivo.", vbExclamation, "atenção"
    Exit Sub
End If

If MsgBox("Deseja criar um nº de documento para estes lançamentos ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
   'gera doc
   Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       nLastDoc = !MAXIMO + 1
      .Close
   End With
   lblNumDoc.Caption = nLastDoc & "-" & RetornaDVNumDoc(nLastDoc)
  'grava doc
   Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,emissor) VALUES("
   Sql = Sql & nLastDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & ",'" & NomeDeLogin & " (ANEXA DOC)" & "')"
   cn.Execute Sql, rdExecDirect
  'grava motivo
   Sql = "INSERT NUMDOCUMENTOMOTIVO (NUMDOCUMENTO,MOTIVO) VALUES("
   Sql = Sql & nLastDoc & ",'" & Mask(txtMotivo.Text) & "')"
   cn.Execute Sql, rdExecDirect
   'GRAVA PARCELADOC
   With grdTemp
        For x = 1 To .Rows - 1
            sAno = .TextMatrix(x, 0)
            sLanc = .TextMatrix(x, 1)
            sSeq = .TextMatrix(x, 2)
            sParc = .TextMatrix(x, 3)
            sComp = .TextMatrix(x, 4)
            Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
            Sql = Sql & Val(sAno) & "," & Val(sLanc) & "," & Val(sSeq) & "," & Val(sParc) & ","
            Sql = Sql & Val(sComp) & "," & nLastDoc & ")"
            cn.Execute Sql, rdExecDirect
        Next
   End With
End If

MsgBox "Documento criado com sucesso.", vbInformation, "Informação"

cmdBaixa.Enabled = False

End Sub

Private Sub cmdCopy_Click()
Clipboard.SetText lblNumDoc.Caption
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Centraliza Me
Me.Top = Me.Top + 1000
CarregaLista
lblDataDoc.Caption = Format(Now, "dd/mm/yyyy")
cmdBaixa.Enabled = True
End Sub

Private Sub CarregaLista()
Dim x As Integer
Dim sSit As String, sVencto As String, sDA As String
Dim sAj As String, nValorPrincipal As Double

With frmDebitoImob.grdExtrato
    nCodReduz = Val(frmDebitoImob.txtCod.Text)
    For x = 1 To .Rows
        If .CellText(x, 12) = "S" Then
           sAno = .CellText(x, 1)
           sLanc = Left$(.CellText(x, 2), 3)
           sSeq = .CellText(x, 3)
           sParc = IIf(.CellText(x, 4) = "Unica", "00", .CellText(x, 4))
           sComp = .CellText(x, 5)
           sSit = Left$(.CellText(x, 6), 2)
           sVencto = .CellText(x, 7)
           sDA = .CellText(x, 8)
           sAj = .CellText(x, 9)
           nValorPrincipal = .CellText(x, 10)
           
           grdTemp.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
             sComp & Chr(9) & sSit & Chr(9) & sVencto & Chr(9) & sDA & Chr(9) & sAj & Chr(9) & _
             FormatNumber(nValorPrincipal, 2)
           
        End If
    Next
    
End With


End Sub


