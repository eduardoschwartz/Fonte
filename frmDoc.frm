VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDoc 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta/Reativação de Documento"
   ClientHeight    =   6300
   ClientLeft      =   2805
   ClientTop       =   3645
   ClientWidth     =   10335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   10335
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      Height          =   1035
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Top             =   5190
      Width           =   10155
   End
   Begin VB.Frame frTemp 
      Height          =   2685
      Left            =   480
      TabIndex        =   22
      Top             =   1050
      Visible         =   0   'False
      Width           =   2805
      Begin MSFlexGridLib.MSFlexGrid grdTemp 
         Height          =   2505
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   4419
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   15658734
         BackColorSel    =   12640511
         ForeColorSel    =   128
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   ">Linha     |^Nº Doc              |^#     "
      End
   End
   Begin VB.Frame frDup 
      BackColor       =   &H00800000&
      Caption         =   "Restituição de Lancamentos Duplicados"
      ForeColor       =   &H00FFFFFF&
      Height          =   3045
      Left            =   3930
      TabIndex        =   18
      Top             =   1050
      Visible         =   0   'False
      Width           =   6195
      Begin MSFlexGridLib.MSFlexGrid grdDup 
         Height          =   1995
         Left            =   270
         TabIndex        =   19
         Top             =   795
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   3519
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   192
         ForeColorFixed  =   16777215
         BackColorSel    =   12640511
         ForeColorSel    =   128
         BackColorBkg    =   16777215
         GridColor       =   16777215
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "<Documento   |^Pagamento  |^Recebimento   |>Valor Pago   |^Banco   |^S  |^#  "
      End
      Begin prjChameleon.chameleonButton cmdFechar 
         Height          =   315
         Left            =   4830
         TabIndex        =   20
         ToolTipText     =   "Retorna Cidadão Selecionado"
         Top             =   390
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Fechar"
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
         MICON           =   "frmDoc.frx":0000
         PICN            =   "frmDoc.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione os documentos que deseja restituir:"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   345
         TabIndex        =   21
         Top             =   570
         Width           =   3270
      End
   End
   Begin VB.TextBox txtNumDoc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   1
      Top             =   90
      Width           =   2025
   End
   Begin MSFlexGridLib.MSFlexGrid grdParc 
      Height          =   3660
      Left            =   30
      TabIndex        =   3
      Top             =   525
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   6456
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      FixedCols       =   0
      BackColorFixed  =   15658734
      BackColorSel    =   12640511
      ForeColorSel    =   128
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDoc.frx":008A
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   9210
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   4560
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDoc.frx":012F
      PICN            =   "frmDoc.frx":014B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdReativa 
      Height          =   345
      Left            =   9075
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Reativar débitos deste documento"
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Restituir"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmDoc.frx":01B9
      PICN            =   "frmDoc.frx":01D5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNF 
      Height          =   345
      Left            =   7785
      TabIndex        =   26
      ToolTipText     =   "Reativar débitos deste documento"
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "N.Fiscais"
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
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmDoc.frx":0248
      PICN            =   "frmDoc.frx":0264
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdView 
      Height          =   315
      Left            =   7740
      TabIndex        =   27
      ToolTipText     =   "Abrir cópia do documento impresso"
      Top             =   4560
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Visualizar"
      ENAB            =   0   'False
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
      MICON           =   "frmDoc.frx":0324
      PICN            =   "frmDoc.frx":0340
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Guia.....:"
      Height          =   255
      Left            =   7740
      TabIndex        =   33
      Top             =   4920
      Width           =   1140
   End
   Begin VB.Label lblValorGuia 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8850
      TabIndex        =   32
      Top             =   4920
      Width           =   1140
   End
   Begin VB.Label lblEmissor 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   840
      TabIndex        =   31
      Top             =   4890
      Width           =   6585
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Emissor..:"
      Height          =   255
      Left            =   90
      TabIndex        =   30
      Top             =   4890
      Width           =   735
   End
   Begin VB.Label lblFile 
      Height          =   285
      Left            =   7020
      TabIndex        =   29
      Top             =   7170
      Width           =   1365
   End
   Begin VB.Label lblPath 
      Height          =   195
      Left            =   5040
      TabIndex        =   28
      Top             =   7215
      Width           =   1545
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "% Desconto......:"
      Height          =   255
      Left            =   7560
      TabIndex        =   25
      Top             =   4290
      Width           =   1290
   End
   Begin VB.Label lblDesconto 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8850
      TabIndex        =   24
      Top             =   4290
      Width           =   1140
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pagamento:"
      Height          =   255
      Left            =   5445
      TabIndex        =   16
      Top             =   4620
      Width           =   1290
   End
   Begin VB.Label lblDataPagto 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "  /  /      "
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6795
      TabIndex        =   15
      Top             =   4620
      Width           =   1140
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Agência.........:"
      Height          =   255
      Left            =   3060
      TabIndex        =   14
      Top             =   4620
      Width           =   1155
   End
   Begin VB.Label lblAgencia 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4260
      TabIndex        =   13
      Top             =   4620
      Width           =   1065
   End
   Begin VB.Label lblBanco 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4260
      TabIndex        =   12
      Top             =   4320
      Width           =   1065
   End
   Begin VB.Label lblValorPago 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6795
      TabIndex        =   11
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Label lblDataDoc 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "  /  /     "
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   4620
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco............:"
      Height          =   255
      Left            =   3060
      TabIndex        =   9
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pago.........:"
      Height          =   255
      Left            =   5445
      TabIndex        =   8
      Top             =   4320
      Width           =   1290
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Documento:"
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   4620
      Width           =   1515
   End
   Begin VB.Label lblValorTaxa 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxa Documento....:"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   4320
      Width           =   1515
   End
   Begin VB.Label lblWait 
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde Carregando..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   4800
      TabIndex        =   2
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o nº do Documento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   2295
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim RdoAux2 As rdoResultset
Dim Sql As String

Private Sub cmdFechar_Click()
Dim x As Integer, Achou As Boolean
With grdDup
    Achou = False
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 5) = "S" Then
            Achou = True
            Exit For
        End If
    Next
End With

If Not Achou Then
    grdParc.TextMatrix(grdParc.Row, 12) = ""
End If

Inicio:
For x = 1 To grdTemp.Rows - 1
    If Val(grdTemp.TextMatrix(x, 0)) = grdParc.Row Then
       If grdTemp.Rows > 2 Then
          grdTemp.RemoveItem (x)
       Else
          grdTemp.Rows = 1
       End If
       GoTo Inicio
    End If
Next

For x = 1 To grdDup.Rows - 1
    If grdDup.TextMatrix(x, 5) = "S" Then
       grdTemp.AddItem grdParc.Row & Chr(9) & grdDup.TextMatrix(x, 0) & Chr(9) & grdDup.TextMatrix(x, 6)
    End If
Next

frDup.Visible = False
grdParc.Enabled = True
cmdReativa.Enabled = True
cmdSair.Enabled = True
grdParc.SetFocus
End Sub


Private Sub cmdNF_Click()

If Val(txtNumDoc.Text) > 0 Then
    frmDocNF.show: frmDocNF.ZOrder 0
    frmDocNF.txtNumDoc.Text = Val(Left(txtNumDoc.Text, 7))
    frmDocNF.cmdConsultar_Click
End If

End Sub

Private Sub cmdReativa_Click()
Dim nNumDoc As Long, nSeqPag As Integer
Dim x As Integer, Y As Integer
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqlanc As Integer
Dim nNumParc As Integer
Dim nCompl As Integer
Dim Achou As Boolean

If NomeDeLogin <> "GLEISE" And NomeDeLogin <> "SCHWARTZ" Then
    Exit Sub
End If
Achou = False
With grdParc
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 12) = "S" Then
           Achou = True
           Exit For
        End If
    Next
End With

If Not Achou Then
   MsgBox "Selecione os lançamentos à reativar.", vbExclamation, "Atenção"
   Exit Sub
End If

Achou = False
With grdParc
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 12) = "" Then
           Achou = True
           Exit For
        End If
    Next
End With

If Achou Then
   If MsgBox("Existem lançamentos não selecionados para reativação." & vbCrLf & "Continuar com a reativação apenas dos lançamentos selecionados ?", vbQuestion + vbYesNo, "Atenção") = vbNo Then
      Exit Sub
   End If
End If

If Not IsDate(lblDataPagto.Caption) Then
   MsgBox "Não foi efetuado baixa para este documento. (" & txtNumDoc.Text & ")" & vbCrLf & "Só é possível reativar um documento previamente baixado.", vbExclamation, "Atenção"
   Exit Sub
End If

'If MsgBox("Deseja executar a restituição destes lançamentos ?" & vbCrLf & vbCrLf & "TODOS OS DADOS REFERENTES AO PAGAMENTO DESTES LANÇAMENTOS SERÃO PERDIDOS. CONTINUAR ?", vbYesNo, "CONFIRMAÇÃO DE REATIVAÇÃO") = vbYes Then
If MsgBox("Deseja executar a restituição destes lançamentos ?", vbYesNo, "CONFIRMAÇÃO DE RESTITUIÇÃO") = vbYes Then
    With grdParc
        For x = 1 To .Rows - 1
            If .TextMatrix(x, 12) = "" Then GoTo Proximo
            nCodReduz = .TextMatrix(x, 1)
            nAnoExercicio = .TextMatrix(x, 0)
            nCodLanc = Val(Left$(.TextMatrix(x, 2), 3))
            nSeqlanc = .TextMatrix(x, 3)
            nNumParc = .TextMatrix(x, 4)
            nCompl = .TextMatrix(x, 5)
            For Y = 1 To grdTemp.Rows - 1
                If grdTemp.TextMatrix(Y, 0) = x Then
                   nNumDoc = Val(Left$(grdTemp.TextMatrix(Y, 1), Len(grdTemp.TextMatrix(Y, 1)) - 1))
                   nSeqPag = Val(grdTemp.TextMatrix(Y, 2))
                  'ATUALIZA A TABELA DEBITOPAGO
                   Sql = "SELECT * FROM DEBITOPAGO "
                   Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                   Sql = Sql & "SEQLANCAMENTO = " & nSeqlanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS NULL"
                   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                   If RdoAux.RowCount > 1 Then
                        Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' "
                        Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                        Sql = Sql & "SEQLANCAMENTO = " & nSeqlanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND SEQPAG=" & nSeqPag
                   Else
                        Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' "
                        Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                        Sql = Sql & "SEQLANCAMENTO = " & nSeqlanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS NULL"
                   End If
                   cn.Execute Sql, rdExecDirect
                   RdoAux.Close
                   Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' "
                   Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                   Sql = Sql & "SEQLANCAMENTO = " & nSeqlanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND SEQPAG=" & nSeqPag
                   cn.Execute Sql, rdExecDirect
                  'ATUALIZA A TABELA NUMDOCUMENTO
'                   Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=0,CODAGENCIA=0,VALORPAGO=0 "
'                   Sql = Sql & "WHERE NUMDOCUMENTO = " & nNumDoc
'                   cn.Execute Sql, rdExecDirect
                End If
            Next
            'SE TODOS OS REGISTROS EM DEBITOPAGO FOREM RESTITUIDOS ENTÃO ATUALIZA DÉBITOPARCELA
'            Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPAGO "
'            Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
'            Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS  NULL"
'            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux
'                If !CONTADOR = 0 Then'
                   'SE FOR ZERO SINAL QUE A PARCELA FOI TOTALMENTE RESTITUIDA
                   'ENTÃO PODEMOS ATUALIZAR O SEU STATUS PARA NÃO PAGO
'                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=3 "
'                    Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
'                    Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl
 '                   cn.Execute Sql, rdExecDirect
'                End If'
'            End With
Proximo:
        Next
    End With
Else
    Exit Sub
End If

MsgBox "Todos os lançamentos descriminados e seus documentos foram reativados.", vbInformation, "INFORMAÇÃO"
txtNumDoc.Text = ""
grdParc.Rows = 1
lblValorTaxa.Caption = ""
lblDataDoc.Caption = "  /  /    "
lblDataPagto.Caption = "  /  /    "
lblValorPago.Caption = 0
lblBanco.Caption = "0"
lblAgencia.Caption = "0"

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdView_Click()
Dim z As Long
If lblFile.Caption <> "" Then
    Call ShellExecute(0&, "open", lblPath.Caption & lblFile.Caption, vbNullString, vbNullString, vbNormalFocus)
End If

End Sub

Private Sub Form_Load()
Centraliza Me
Ocupado
lblWait.Visible = False
Liberado

End Sub

Private Sub grdDup_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
    
   With grdDup
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, 5) = "" Then
           .TextMatrix(.Row, 5) = "S"
           For x = 0 To 6
             .col = x
             .CellBackColor = vbRed
             .CellForeColor = Branco
           Next
        Else
           .TextMatrix(.Row, 5) = ""
           For x = 0 To 6
             .col = x
             .CellBackColor = Branco
             .CellForeColor = vbBlack
           Next
        End If
        .col = 0
        .ColSel = 6
   End With
End If

End Sub

Private Sub grdParc_KeyPress(KeyAscii As Integer)
Dim nNumDoc As Long
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqlanc As Integer
Dim nNumParc As Integer
Dim nCompl As Integer

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   With grdParc
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, 12) = "" Then
            nCodReduz = .TextMatrix(.Row, 1)
            nAnoExercicio = .TextMatrix(.Row, 0)
            nCodLanc = Val(Left$(.TextMatrix(.Row, 2), 3))
            nSeqlanc = .TextMatrix(.Row, 3)
            nNumParc = .TextMatrix(.Row, 4)
            nCompl = .TextMatrix(.Row, 5)
            Sql = "SELECT CODREDUZIDO, ANOEXERCICIO, CODLANCAMENTO, SEQLANCAMENTO, NUMPARCELA, CODCOMPLEMENTO,"
            Sql = Sql & "SEQPAG, DATAPAGAMENTO, DATARECEBIMENTO,VALORPAGO, CODBANCO, CODAGENCIA, RESTITUIDO, NumDocumento "
            Sql = Sql & "From DEBITOPAGO WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
            Sql = Sql & "SEQLANCAMENTO = " & nSeqlanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS NULL order by anoexercicio,codlancamento,numparcela"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    grdParc.TextMatrix(grdParc.Row, 12) = "S"
                    For x = 0 To 12
                      grdParc.col = x
                      grdParc.CellBackColor = vbRed
                      grdParc.CellForeColor = Branco
                    Next
                Else
                    MsgBox "Este lançamento ja foi restituido.", vbExclamation, "Atenção"
                End If
            End With
        Else
           .TextMatrix(.Row, 12) = ""
           For x = 0 To 12
             .col = x
             .CellBackColor = Branco
             .CellForeColor = vbBlack
           Next
        End If
        .col = 0
        .ColSel = 12
        If .TextMatrix(.Row, 11) = "N" And .TextMatrix(.Row, 12) = "S" Then
           grdTemp.AddItem .Row & Chr(9) & txtNumDoc.Text & Chr(9) & "0"
        ElseIf .TextMatrix(.Row, 11) = "N" And .TextMatrix(.Row, 12) = "" Then
Inicio:
           For x = 1 To grdTemp.Rows - 1
               If Val(grdTemp.TextMatrix(x, 0)) = grdParc.Row Then
                  If grdTemp.Rows > 2 Then
                     grdTemp.RemoveItem (x)
                  Else
                     grdTemp.Rows = 1
                  End If
                  GoTo Inicio
               End If
           Next
        ElseIf .TextMatrix(.Row, 11) = "S" And .TextMatrix(.Row, 12) = "" Then
Inicio2:
           For x = 1 To grdTemp.Rows - 1
               If Val(grdTemp.TextMatrix(x, 0)) = grdParc.Row Then
                  If grdTemp.Rows > 2 Then
                     grdTemp.RemoveItem (x)
                  Else
                     grdTemp.Rows = 1
                  End If
                  GoTo Inicio2
               End If
           Next
        End If
        
        If .TextMatrix(.Row, 11) = "S" And .TextMatrix(.Row, 12) = "S" Then
           frDup.Visible = True
           grdParc.Enabled = False
           cmdReativa.Enabled = False
           cmdSair.Enabled = False
           grdDup.Rows = 1
           grdDup.SetFocus
           nNumDoc = Val(Left$(txtNumDoc.Text, Len(txtNumDoc.Text) - 1))
           
           nCodReduz = .TextMatrix(.Row, 1)
           nAnoExercicio = .TextMatrix(.Row, 0)
           nCodLanc = Val(Left$(.TextMatrix(.Row, 2), 3))
           nSeqlanc = .TextMatrix(.Row, 3)
           nNumParc = .TextMatrix(.Row, 4)
           nCompl = .TextMatrix(.Row, 5)
            
           Sql = "SELECT CODREDUZIDO,VALORPAGO,DATAPAGAMENTO,DATARECEBIMENTO,NUMDOCUMENTO,CODBANCO,SEQPAG  FROM DEBITOPAGO WHERE "
           Sql = Sql & "CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc
           Sql = Sql & " AND NUMPARCELA=" & nNumParc & " AND SEQLANCAMENTO=" & nSeqlanc & " AND CODCOMPLEMENTO=" & nCompl & " AND RESTITUIDO IS  NULL"
           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux2
                grdDup.Rows = 1
                Do Until .EOF
                   If Not IsNull(!NumDocumento) Then
                      grdDup.AddItem !NumDocumento & "-" & RetornaDVNumDoc(!NumDocumento) & Chr(9) & Format(!DataPagamento, "dd/mm/yyyy") & Chr(9) & Format(!datarecebimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(!ValorPago + CDbl(lblValorTaxa.Caption), 2) & Chr(9) & !CodBanco & Chr(9) & "" & Chr(9) & !SEQPAG
                   End If
                  .MoveNext
                Loop
               .Close
           End With
        End If
   End With
End If

End Sub


Public Sub txtNumDoc_KeyPress(KeyAscii As Integer)
Dim sNumDoc As String, sDataDoc As String, sFile As String



If cGetInputState() <> 0 Then DoEvents
'lblWait.Visible = True
'lblWait.Refresh
grdParc.Rows = 1
grdDup.Rows = 1
grdTemp.Rows = 1
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    CarregaDoc
 '   If grdParc.Rows > 1 Then
 '       sNumDoc = Left(txtNumDoc.Text, Len(txtNumDoc.Text) - 1)
 '       sDataDoc = lblDataDoc.Caption
'        sFile = "\\192.168.200.130\atualizagti\segundavia\" & Mid(sDataDoc, 4, 2) & Mid(sDataDoc, 7, 4) & "\" & Format(CLng(sNumDoc), "000000000") & "*.pdf"
'        lblPath.Caption = "\\192.168.200.130\atualizagti\segundavia\" & Mid(sDataDoc, 4, 2) & Mid(sDataDoc, 7, 4) & "\"
'        lblFile.Caption = Dir(sFile)
'        If lblFile.Caption <> "" Then
'            Call ShellExecute(0&, vbNullString, Dir(sFile), vbNullString, vbNullString, vbNormalFocus)
'            cmdView.Enabled = True
'        Else
'            cmdView.Enabled = False
'        End If
 '   Else
  '      cmdView.Enabled = False
  '  End If
ElseIf Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
   KeyAscii = 0
End If

End Sub

Private Sub CarregaDoc()
Dim sVencto As String
Dim sDataBase As String
Dim nNumDoc As Long
Dim nValorPrincipal As Double
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim nStatus As Integer, sStatus As String
Dim nSomaTotal As Double, nPerc As Double
Dim nSomaPrincipal As Double, RdoTmp As rdoResultset
Dim bDupl As Boolean, qd As New rdoQuery
Dim x As Integer, nSomaL As Double, nSomaJ As Double, nSomaM As Double, nSomaC As Double, bAchou As Boolean
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeqlanc As Integer, nParc As Integer, nCompl As Integer

On Error GoTo Erro
If txtNumDoc.Text = "" Then Exit Sub

nNumDoc = Val(Left$(txtNumDoc.Text, Len(txtNumDoc.Text) - 1))
'VALIDA DIGITO VERIFICADOR
If Val(Right$(txtNumDoc.Text, 1)) <> RetornaDVNumDoc(nNumDoc) Then
   MsgBox "Digito Verificador Inválido", vbExclamation, "Atenção"
   GoTo fim
End If
'VERIFICA SE O DOCUMENTO EXISTE NA TABELA NUMDOCUMENTO
Sql = "SELECT NUMDOCUMENTO FROM NUMDOCUMENTO "
Sql = Sql & "WHERE NUMDOCUMENTO=" & nNumDoc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
   If .RowCount = 0 Then
        lblValorTaxa.Caption = ""
        lblDataDoc.Caption = "  /  /    "
        lblDataPagto.Caption = "  /  /    "
        lblValorPago.Caption = 0
        lblBanco.Caption = "0"
        lblAgencia.Caption = "0"
        lblDesconto.Caption = "0,00%"
        lblEmissor.Caption = ""
       .Close
        MsgBox "Nº de Documento não encontrado.", vbExclamation, "Atenção"
        GoTo fim
   End If
   'CARREGA OS DEBITOS DESTE DOCUMENTO
'   Sql = "SELECT PARCELADOCUMENTO.CODREDUZIDO,PARCELADOCUMENTO.ANOEXERCICIO,PARCELADOCUMENTO.CODLANCAMENTO,LANCAMENTO.DESCREDUZ,PARCELADOCUMENTO.SEQLANCAMENTO,"
'   Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA,PARCELADOCUMENTO.CODCOMPLEMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,NUMDOCUMENTO.DATADOCUMENTO,NUMDOCUMENTO.CODBANCO,NUMDOCUMENTO.VALORTAXADOC,NUMDOCUMENTO.PERCISENCAO,"
'   Sql = Sql & "NUMDOCUMENTO.CODAGENCIA,NUMDOCUMENTO.EMISSOR,NUMDOCUMENTO.VALORGUIA,NUMDOCUMENTO.VALORPAGO,DEBITOPARCELA.STATUSLANC,SITUACAOLANCAMENTO.DESCSITUACAO,DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE "
'   Sql = Sql & "FROM PARCELADOCUMENTO INNER JOIN DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
 '  Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND PARCELADOCUMENTO.NumParcela = DEBITOPARCELA.NumParcela AND "
 '  Sql = Sql & "PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO Inner Join LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO Inner Join SITUACAOLANCAMENTO ON "
'   Sql = Sql & "DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO Inner Join NUMDOCUMENTO ON PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO Where PARCELADOCUMENTO.NumDocumento = " & nNumDoc
 Sql = "SELECT parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, lancamento.descreduz, parceladocumento.seqlancamento, "
Sql = Sql & "parceladocumento.numparcela, parceladocumento.codcomplemento, parceladocumento.numdocumento, numdocumento.datadocumento, numdocumento.codbanco,"
Sql = Sql & "numdocumento.valortaxadoc, numdocumento.percisencao, numdocumento.codagencia,numdocumento.emissor ,numdocumento.valorguia, numdocumento.valorpago,"
Sql = Sql & "debitoparcela.statuslanc, situacaolancamento.descsituacao, debitoparcela.datavencimento, debitoparcela.datadebase, numdocumento.userid,Usuario.NomeLogin "
Sql = Sql & "FROM parceladocumento INNER JOIN "
Sql = Sql & "debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND "
Sql = Sql & "parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
Sql = Sql & "lancamento ON debitoparcela.codlancamento = lancamento.codlancamento INNER JOIN "
Sql = Sql & "situacaolancamento ON debitoparcela.statuslanc = situacaolancamento.codsituacao INNER JOIN "
Sql = Sql & "numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento LEFT OUTER JOIN "
Sql = Sql & "usuario ON numdocumento.userid = usuario.Id Where PARCELADOCUMENTO.NumDocumento = " & nNumDoc
  
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount = 0 Then
          MsgBox "Lançamentos não encontrados para este documento.", vbExclamation, "Atenção"
       Else
       nCodReduz = !CODREDUZIDO
       nAno = !AnoExercicio
       nLanc = !CodLancamento
       nSeqlanc = !SeqLancamento
       nParc = !NumParcela
       nCompl = !CODCOMPLEMENTO
       nStatus = !statuslanc
       sStatus = !DescSituacao
       lblBanco.Caption = SubNull(!CodBanco)
       lblAgencia.Caption = SubNull(!CodAgencia)
       lblValorPago.Caption = FormatNumber(!ValorPago, 2)
       If IsNull(!VALORGUIA) Then
          lblValorGuia.Caption = FormatNumber(0, 2)
       Else
          lblValorGuia.Caption = FormatNumber(!VALORGUIA, 2)
       End If
       lblDataDoc.Caption = IIf(IsNull(!DATADOCUMENTO), "  /  /   ", Format(!DATADOCUMENTO, "dd/mm/yyyy"))
       lblDesconto.Caption = SubNull(!percisencao) & "%"
       lblEmissor.Caption = IIf(SubNull(!NomeLogin) = "", SubNull(!emissor), !NomeLogin)
       'SE NÃO TIVER TAXADOC SINAL QUE VEIO DA SMARK ENTÃO PEGAMOS A TAXADOC DO 1º LANCAMENTO
       If IsNull(!ValorTaxaDoc) Or !ValorTaxaDoc = 0 Then
          Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
          Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          With RdoAux2
              If .RowCount > 0 Then
                 lblValorTaxa.Caption = FormatNumber(!ValorTributo, 2)
              Else
                 lblValorTaxa.Caption = "0,00"
              End If
             .Close
          End With
       Else
          lblValorTaxa.Caption = FormatNumber(!ValorTaxaDoc, 2)
       End If
       'DATA DE PAGAMENTO
       Sql = "SELECT CODREDUZIDO,DATAPAGAMENTO,CODBANCO FROM DEBITOPAGO "
       Sql = Sql & "Where NumDocumento = " & nNumDoc
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
          If .RowCount > 0 Then
             lblDataPagto.Caption = IIf(IsNull(!DataPagamento) Or !DataPagamento = CDate("01/01/1900"), "  /  /   ", Format(!DataPagamento, "dd/mm/yyyy"))
          End If
         .Close
       End With

        
            
        Do Until .EOF
            sVencto = Format(!DataVencimento, "dd/mm/yyyy")
            sDataBase = Format(!DATADEBASE, "dd/mm/yyyy")
            nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValorTotal = 0
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoTmp.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = !CODREDUZIDO
            qd(1) = !CODREDUZIDO
            qd(2) = !AnoExercicio
            qd(3) = !AnoExercicio
            qd(4) = !CodLancamento
            qd(5) = !CodLancamento
            qd(6) = !SeqLancamento
            qd(7) = !SeqLancamento
            qd(8) = !NumParcela
            qd(9) = !NumParcela
            qd(10) = !CODCOMPLEMENTO
            qd(11) = !CODCOMPLEMENTO
            qd(12) = 1
            qd(13) = 99
            qd(14) = IIf(IsDate(lblDataPagto), Format(lblDataPagto, "mm/dd/yyyy"), Format(Now, "mm/dd/yyyy"))
            qd(15) = NomeDoUsuario
            Set RdoTmp = qd.OpenResultset(rdOpenKeyset)
            With RdoTmp
                Do Until .EOF
                    nValorPrincipal = nValorPrincipal + !ValorTributo
                    nValorJuros = nValorJuros + !ValorJuros
                    nValorMulta = nValorMulta + !ValorMulta
                    nValorCorrecao = nValorCorrecao + !ValorCorrecao
                    nValorTotal = nValorTotal + !ValorTotal
                   .MoveNext
                Loop
               .Close
            End With
        
            If lblDesconto.Caption = "%" Then
               nPerc = 0
            Else
                nPerc = RetornaNumero(lblDesconto.Caption)
            End If
     '       If nPerc > 0 And IsDate(lblDataPagto.Caption) Then
     '           nValorCorrecao = nValorCorrecao - FormatNumber(nValorCorrecao * nPerc / 100, 2)
     '           nValorJuros = nValorJuros - FormatNumber(nValorJuros * nPerc / 100, 2)
     '           nValorMulta = nValorMulta - FormatNumber(nValorMulta * nPerc / 100, 2)
     '           nValorTotal = nValorPrincipal + nValorJuros + nValorMulta + nValorCorrecao
     '       End If


            
            Sql = "SELECT CODREDUZIDO FROM DEBITOPAGO "
            Sql = Sql & "WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
            Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND RESTITUIDO IS  NULL"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                  If .RowCount > 1 Then
                     bDupl = True
                  Else
                     bDupl = False
                  End If
                 .Close
            End With
            
            grdParc.AddItem !AnoExercicio & Chr(9) & Format(!CODREDUZIDO, "000000") & Chr(9) & Format(!CodLancamento, "000") & " - " & !descreduz & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & Format(!NumParcela, "00") & Chr(9) & _
              !CODCOMPLEMENTO & Chr(9) & FormatNumber(nValorPrincipal, 2) & Chr(9) & FormatNumber(nValorMulta, 2) & Chr(9) & _
              FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2) & Chr(9) & FormatNumber(nValorTotal, 2) & Chr(9) & IIf(bDupl, "S", "N")
          .MoveNext
       Loop
       End If
      .Close
      
   End With
End With

txtObs.Text = ""
Sql = "SELECT obsparcela.*, usuario.nomelogin FROM obsparcela INNER JOIN usuario ON obsparcela.userid = usuario.Id WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqlanc & " AND NUMPARCELA=" & nParc
Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        txtObs.Text = txtObs.Text & !obs & vbCrLf
       .MoveNext
    Loop
   .Close
End With



fim:
lblWait.Visible = False
lblWait.Refresh

Exit Sub

Erro:
If rdoErrors(0).Number = 8115 Then
   MsgBox "Nº de Documento inválido.", vbExclamation, "Atenção"
Else
   MsgBox Err.Description
End If
Resume Next

End Sub
