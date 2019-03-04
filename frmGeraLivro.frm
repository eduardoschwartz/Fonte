VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmGeraLivro 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração dos Livros da Divida Ativa"
   ClientHeight    =   4980
   ClientLeft      =   6345
   ClientTop       =   3375
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   8070
   Begin VB.CheckBox chkAj 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Ajuizado"
      Height          =   240
      Left            =   2565
      TabIndex        =   18
      Top             =   4635
      Width           =   1050
   End
   Begin VB.CheckBox chkReparc 
      Caption         =   "Reparcelamento"
      Height          =   225
      Left            =   150
      TabIndex        =   10
      Top             =   4680
      Width           =   2025
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Todos"
      Height          =   195
      Left            =   1425
      TabIndex        =   16
      Top             =   3180
      Width           =   765
   End
   Begin VB.ListBox lstStatus 
      Appearance      =   0  'Flat
      Height          =   1155
      Left            =   90
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   15
      Top             =   3435
      Width           =   2100
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Sintético"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   5220
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Analítico"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   5460
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ListBox lstTipo 
      Appearance      =   0  'Flat
      Height          =   1155
      Left            =   90
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   1935
      Width           =   2100
   End
   Begin VB.CheckBox chkT 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Todos"
      Height          =   195
      Left            =   1425
      TabIndex        =   11
      Top             =   1680
      Width           =   765
   End
   Begin VB.ListBox lstAno 
      Appearance      =   0  'Flat
      Height          =   1155
      ItemData        =   "frmGeraLivro.frx":0000
      Left            =   135
      List            =   "frmGeraLivro.frx":0002
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   435
      Width           =   2100
   End
   Begin VB.CheckBox chkA 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Todos"
      Height          =   270
      Left            =   1470
      TabIndex        =   8
      Top             =   120
      Width           =   765
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   3450
      Left            =   2490
      TabIndex        =   7
      Top             =   420
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   6085
      _Version        =   393216
      Rows            =   1
      Cols            =   5
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
      FormatString    =   "^Ano       |<Tipo de Livro       |^Nº do Livro   |^Nº Antigo  |>Valor Total          "
   End
   Begin MSComctlLib.ProgressBar Pb2 
      Height          =   225
      Left            =   2940
      TabIndex        =   4
      Top             =   3990
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   225
      Left            =   2940
      TabIndex        =   1
      Top             =   4350
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   360
      Left            =   6435
      TabIndex        =   0
      ToolTipText     =   "Cancelar Edição"
      Top             =   4035
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Emitir Livro(s)"
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
      MICON           =   "frmGeraLivro.frx":0004
      PICN            =   "frmGeraLivro.frx":0020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   6420
      TabIndex        =   17
      ToolTipText     =   "Cancelar Edição"
      Top             =   4530
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Cancelados"
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
      MICON           =   "frmGeraLivro.frx":00BF
      PICN            =   "frmGeraLivro.frx":00DB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskData 
      Height          =   285
      Left            =   4455
      TabIndex        =   20
      Top             =   4635
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BackColor       =   16777215
      MouseIcon       =   "frmGeraLivro.frx":0235
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Data..:"
      Height          =   195
      Left            =   3825
      TabIndex        =   19
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label lblPB2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5400
      TabIndex        =   6
      Top             =   4005
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2610
      TabIndex        =   5
      Top             =   3990
      Width           =   270
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2610
      TabIndex        =   3
      Top             =   4350
      Width           =   270
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5400
      TabIndex        =   2
      Top             =   4395
      Width           =   495
   End
   Begin VB.Menu mnuTipo 
      Caption         =   "Tipo"
      Visible         =   0   'False
      Begin VB.Menu mnuImovel 
         Caption         =   "Imóvel"
      End
      Begin VB.Menu mnuEmpresa 
         Caption         =   "Empresa"
      End
      Begin VB.Menu mnuCidadao 
         Caption         =   "Cidadão"
      End
   End
End
Attribute VB_Name = "frmGeraLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSituacao As Integer
    sSituacao As String
    sVencto As String
    sDA As String
    sAj As String
    nCodTributo As Double
    nValorTributo As Double
    sDescTributo As String
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorAtual As Double
    nValorGeral As Double
    nValorHon As Double
    nValorJurApl As Double
    nSaldo As Double
    nCodBanco As Integer
    dDataPag As Date
    sNome As String
    sFullLanc As String
    nPagina As Integer
    dDataInscricao As Date
    nCertidao As Integer
    sNumProcesso As String
End Type

Private Type DIVIDATRIBUTO
    nNumLivro As Integer
    nCodTributo As Integer
    sDescTributo As String
    nSomaP As Double
    nSomaM As Double
    nSomaJ As Double
    nSomaC As Double
    nSomaT As Double
End Type
Dim xImovel As clsImovel

Private Sub chkA_Click()
Dim x As Integer

If Not bExec Then Exit Sub
For x = 0 To lstAno.ListCount - 1
   lstAno.Selected(x) = IIf(chkA.value, 1, 0)
Next

End Sub

Private Sub chkT_Click()
Dim x As Integer

If Not bExec Then Exit Sub
For x = 0 To lstTipo.ListCount - 1
   lstTipo.Selected(x) = IIf(chkT.value, 1, 0)
Next

End Sub

Private Sub cmdCancel_Click()
PopupMenu mnuTipo
End Sub

Private Sub cmdExec_Click()
Dim sTipo As String
If Opt(0).value = False And Opt(1).value = False Then
    MsgBox "Selecione o tipo do Relatório.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskData.Text) Then
    MsgBox "Data inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

For x = 0 To lstAno.ListCount - 1
    If lstAno.Selected(x) = True Then
        lstAno.ListIndex = x
        sAno = sAno & lstAno.Text & ","
    End If
Next
If sAno = "" Then
    MsgBox "Selecione ao menos um ano de exercício.", vbCritical, "Erro"
    Exit Sub
End If


For x = 0 To lstTipo.ListCount - 1
    If lstTipo.Selected(x) = True Then
        lstTipo.ListIndex = x
        Y = x + 1
        sTipo = sTipo & lstTipo.ItemData(lstTipo.ListIndex) & ","
    End If
Next
sTipo = Chomp(sTipo, chomp_righT, 1)
If sTipo = "" Then
    MsgBox "Selecione ao menos um tipo de livro.", vbCritical, "Erro"
    Exit Sub
End If



Ocupado
cmdExec.Enabled = False

CarregaGrid
Me.Refresh
If cGetInputState() <> 0 Then DoEvents
Calcula
'Calculo2
If Opt(0).value = True Then
    frmReport.ShowReport "DIVIDATIVATOTAL", frmMdi.HWND, Me.HWND
Else
    If chkReparc.value = vbUnchecked Then
        frmReport.ShowReport "DIVIDATIVA", frmMdi.HWND, Me.HWND
    Else
        If lstTipo.ListIndex = 0 Then
           frmReport.ShowReport "DIVIDATIVAPARCIPTU", frmMdi.HWND, Me.HWND
        ElseIf lstTipo.ListIndex = 1 Then
           frmReport.ShowReport "DIVIDATIVAPARCISS", frmMdi.HWND, Me.HWND
        Else
            frmReport.ShowReport "DIVIDATIVAPARCNAOTRIBUTAVEL", frmMdi.HWND, Me.HWND
        End If
    End If
End If

Sql = "delete  from DIVIDATIVA"
cn.Execute Sql, rdExecDirect
Sql = "delete  from DIVIDATIVATOTAL"
cn.Execute Sql, rdExecDirect

Liberado
cmdExec.Enabled = True
End Sub

Private Sub Form_Load()
Centraliza Me
Set xImovel = New clsImovel
bExec = True
CarregaAno
CarregaTipo
CarregaStatus
mskData.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub lstAno_ItemCheck(Item As Integer)

bExec = False
If lstAno.Selected(Item) = False Then chkA.value = 0
bExec = True

End Sub

Private Sub lstTipo_ItemCheck(Item As Integer)

bExec = False
If lstTipo.Selected(Item) = False Then chkT.value = 0
bExec = True

End Sub

Private Sub CarregaAno()
Dim x As Integer
For x = 1990 To 2027
    lstAno.AddItem x
Next

End Sub

Private Sub CarregaTipo()
Sql = "SELECT CODTIPO,DESCTIPO FROM TIPOLIVRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       lstTipo.AddItem !DESCTIPO
       lstTipo.ItemData(lstTipo.NewIndex) = !CodTipo
      .MoveNext
    Loop
End With

End Sub

Private Sub CarregaStatus()
Sql = "SELECT CODSITUACAO,DESCSITUACAO FROM SITUACAOLANCAMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       lstStatus.AddItem !DescSituacao & " (" & Format(!codsituacao, "00") & ")"
       lstStatus.ItemData(lstStatus.NewIndex) = !codsituacao
      .MoveNext
    Loop
End With

End Sub

Private Sub CarregaGrid()

Dim sAno As String, sTipo As String, x As Integer, Y As Integer

grdTemp.Rows = 1

For x = 0 To lstAno.ListCount - 1
    If lstAno.Selected(x) = True Then
        lstAno.ListIndex = x
        sAno = sAno & lstAno.Text & ","
    End If
Next
sAno = Chomp(sAno, chomp_righT, 1)

For x = 0 To lstTipo.ListCount - 1
    If lstTipo.Selected(x) = True Then
        lstTipo.ListIndex = x
        Y = x + 1
        sTipo = sTipo & lstTipo.ItemData(lstTipo.ListIndex) & ","
    End If
Next
sTipo = Chomp(sTipo, chomp_righT, 1)

Sql = "SELECT ANO,CODTIPO, NUMERO,DATAABERTURA,DATAENCERRAMENTO FROM LIVRO WHERE ANO in (" & sAno & ")"
Sql = Sql & " AND CODTIPO in (" & sTipo & ")"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !CodTipo = 1 Then
           sTipo = CStr(!CodTipo) & " - IPTU"
        ElseIf !CodTipo = 2 Then
           sTipo = CStr(!CodTipo) & " - ISSQN"
        ElseIf !CodTipo = 3 Then
           sTipo = CStr(!CodTipo) & " - TAXA DIV./NÃO TRIBUTÁVEL"
        End If
        Sql = "SELECT NUMEROOLD FROM GRADELIVRO WHERE NUMEROLIVRO=" & !Numero & " ORDER BY ANO,NUMEROOLD"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                Do Until .EOF
                    grdTemp.AddItem RdoAux!Ano & Chr(9) & sTipo & Chr(9) & RdoAux!Numero & Chr(9) & !NUMEROOLD
                   .MoveNext
                Loop
            Else
                grdTemp.AddItem RdoAux!Ano & Chr(9) & sTipo & Chr(9) & RdoAux!Numero & Chr(9) & RdoAux!Numero
            End If
           .Close
        End With
       .MoveNext
    Loop
   .Close
End With

If grdTemp.Rows = 1 Then
    grdTemp.AddItem lstAno.Text & Chr(9) & Y & Chr(9) & 100 & Chr(9) & 100
End If


End Sub

Private Sub Calcula()
Dim nAno As Integer, nCodTipo As Integer, nNumLivro As Integer, z As Integer, sNumProcesso As String
Dim RdoC As rdoResultset, xId As Long, nNumRec As Long, nAnoLinha As Integer, RdoAux2 As rdoResultset
Dim nValorCorrecao As Double, nValorTotal As Double, sDescTributo As String
Dim sProp As String, sStatus As String, sInscricao As String, sEndereco As String
Dim sCompl As String, sCep As String, sBairro As String, sCidade As String, sUF As String
Dim nValorJuros As Double, nValorMulta As Double, dDataInsc As Date, bJuros As Boolean, bMulta As Boolean
Dim aSomaTributo() As DIVIDATRIBUTO, bFind As Boolean
ReDim aSomaTributo(0)

sStatus = ""
For x = 0 To lstStatus.ListCount - 1
    If lstStatus.Selected(x) = True Then
        sStatus = sStatus & CStr(lstStatus.ItemData(x)) & ","
    End If
Next
If sStatus <> "" Then
    sStatus = Left(sStatus, Len(sStatus) - 1)
    sStatus = " AND STATUSLANC in (" & sStatus & ")"
End If


Sql = "delete  from DIVIDATIVA"
cn.Execute Sql, rdExecDirect
Sql = "delete  from DIVIDATIVATOTAL"
cn.Execute Sql, rdExecDirect

nValorTotal = 0
xId = 0
With grdTemp
    For z = 1 To .Rows - 1
        .TopRow = z
        If chkReparc.value = 1 Then
            If z > 1 Then
                nAnoLinha = .TextMatrix(z, 0)
                If nAnoLinha = nAno Then
                    GoTo PROXIMO2
                End If
            End If
        End If
        
        nAno = .TextMatrix(z, 0)
        nCodTipo = Val(Left$(.TextMatrix(z, 1), 1))
        nNumLivro = .TextMatrix(z, 3)
        xId = 0
        
        If nAno < 2004 Then
            If nCodTipo = 1 Then
                If chkReparc.value = 1 Then
                    Sql = "SELECT * FROM VWDIVIDAATIVAR WHERE CODREDUZIDO < 100000 AND ANOEXERCICIO=" & nAno & " AND CODTRIBUTO<>3 AND CODLANCAMENTO=20" & sStatus
                Else
                    Sql = "SELECT * FROM VWDIVIDAATIVA WHERE NUMEROLIVRO=" & nNumLivro & " AND CODREDUZIDO < 100000 AND YEAR(DATAINSCRICAO)=" & nAno & " AND (CODTRIBUTO=1 or codtributo=2) " & sStatus
                End If
            ElseIf nCodTipo = 2 Then
                If chkReparc.value = 1 Then
                    Sql = "SELECT * FROM VWDIVIDAATIVAR WHERE (CODREDUZIDO > 100000 AND CODREDUZIDO<500000)  AND ANOEXERCICIO=" & nAno & " AND CODTRIBUTO<>3 AND VALORTRIBUTO>0 and CODLANCAMENTO=20" & sStatus
                Else
                    Sql = "SELECT * FROM VWDIVIDAATIVA WHERE NUMEROLIVRO=" & nNumLivro & " AND (CODREDUZIDO > 100000) AND YEAR(DATAINSCRICAO)=" & nAno & " AND VALORTRIBUTO>0 and CODtributo in (11,12,13,14,25,179,180,181,182,183,502)" & sStatus
                End If
            ElseIf nCodTipo = 3 Then
                If chkReparc.value = 1 Then
                    Sql = "SELECT * FROM VWDIVIDAATIVAR WHERE CODREDUZIDO > 500000  AND ANOEXERCICIO=" & nAno & " AND CODTRIBUTO<>3 AND CODLANCAMENTO=20" & sStatus
                Else
                    Sql = "SELECT * FROM VWDIVIDAATIVA WHERE NUMEROLIVRO=" & nNumLivro & " AND YEAR(DATAINSCRICAO)=" & nAno & " AND CODTRIBUTO<>3" & sStatus
                End If
            End If
        Else
            If nCodTipo = 1 Then
                If chkReparc.value = 1 Then
                    Sql = "SELECT * FROM VWDIVIDAATIVAR WHERE  CODREDUZIDO < 100000  AND ANOEXERCICIO=" & nAno & " AND CODTRIBUTO<>3 AND CODLANCAMENTO=20" & sStatus
                Else
                    Sql = "SELECT * FROM VWDIVIDAATIVA WHERE YEAR(DATAINSCRICAO)=" & nAno & " AND CODREDUZIDO < 100000 AND NUMEROLIVRO=" & nNumLivro & " AND CODTRIBUTO<>3 AND (CODLANCAMENTO=1 or CODLANCAMENTO=29)" & sStatus
                End If
            ElseIf nCodTipo = 2 Then
                If chkReparc.value = 1 Then
                    Sql = "SELECT * FROM VWDIVIDAATIVAR WHERE CODREDUZIDO>=100000 AND ANOEXERCICIO=" & nAno & " AND  CODTRIBUTO<>3 AND VALORTRIBUTO>0 and CODLANCAMENTO=20 " & sStatus
                Else
                    Sql = "SELECT * FROM VWDIVIDAATIVA WHERE NUMEROLIVRO=" & nNumLivro & " AND (CODREDUZIDO >= 100000) AND YEAR(DATAINSCRICAO)=" & nAno & " AND CODTRIBUTO in (11,12,13,14,19,25,179,180,181,182,183,184,502) AND VALORTRIBUTO>0" & sStatus
                End If
            ElseIf nCodTipo = 3 Or nCodTipo = 4 Or nCodTipo = 5 Then
                If chkReparc.value = 1 Then
                    Sql = "SELECT * FROM VWDIVIDAATIVAR WHERE  CODREDUZIDO >= 100000  AND ANOEXERCICIO =" & nAno & " AND CODTRIBUTO<>3 AND CODLANCAMENTO=20" & sStatus
                Else
                    Sql = "SELECT * FROM VWDIVIDAATIVA WHERE YEAR(DATAINSCRICAO)=" & nAno & " AND NUMEROLIVRO=" & nNumLivro & " AND CODLANCAMENTO<>20 AND CODTRIBUTO NOT in (1,2,3,11,12,13,14,19,25,179,180,181,182,183,184,502)  " & sStatus
                End If
            End If
        End If
        If chkAj.value = vbChecked Then
            Sql = Sql & " AND DATAAJUIZA IS NOT NULL"
        End If
        
        Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoC
        
            nNumRec = .RowCount
            If nNumRec = 0 Then GoTo PROXIMO2
            If Not IsNull(!datainscricao) Then
                dDataInsc = Format(!datainscricao, "dd/mm/yyyy")
            Else
                dDataInsc = Format(Now, "dd/mm/yyyy")
            End If
            Do Until .EOF
'                If !CODREDUZIDO = 2610 Then MsgBox "teste"
               If xId Mod 80 = 0 Then
                  CallPb2 xId, nNumRec
               End If

                sProp = ""
                '**** Endereço ******
                With xImovel
                    If RdoC!CODREDUZIDO < 100000 Then
                       .RetornaEndereco RdoC!CODREDUZIDO, Imobiliario, Localizacao
                        Sql = "SELECT * FROM vwfullimovel2 WHERE CODREDUZIDO=" & RdoC!CODREDUZIDO
                        Set rdoc2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                        If rdoc2.RowCount > 0 Then
                            sProp = rdoc2!NomeCidadao
                            rdoc2.Close
                        End If
                    ElseIf RdoC!CODREDUZIDO >= 100000 And RdoC!CODREDUZIDO < 500000 Then
                       .RetornaEndereco RdoC!CODREDUZIDO, Mobiliario, Localizacao
                        Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & RdoC!CODREDUZIDO
                        Set rdoc2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                        If rdoc2.RowCount > 0 Then
                            sProp = rdoc2!RazaoSocial
                            rdoc2.Close
                        End If
                    Else
                        Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & RdoC!CODREDUZIDO
                        Set rdoc2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                        If rdoc2.RowCount > 0 Then
                            sProp = rdoc2!NomeCidadao
                            rdoc2.Close
                        End If
                       .RetornaEndereco RdoC!CODREDUZIDO, cidadao, cadastrocidadao
                    End If
                
                    sEndereco = .Endereco & ", " & .Numero
                    sCompl = .Complemento
                    sCep = Format(.Cep, "00000-000")
                    sBairro = .Bairro
                    sCidade = .Cidade
                    sUF = .UF
                End With
                
                  
               sInscricao = ""
               'If DateDiff("y", !DataVencimento, CDate("31/12/2013")) > 0 Then
         '      If Year(!DataVencimento) < 2014 Then
                   nValorCorrecao = CalculaCorrecao(!ValorTributo, Format(!DataVencimento, "dd/mm/yyyy"), mskData.Text)
         '      Else
         '          nValorCorrecao = 0
         '      End If
               
               Sql = "SELECT DESCTRIBUTO,MULTA,JUROS FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                        sDescTributo = !desctributo
                        bJuros = !Juros
                        bMulta = !Multa
                   Else
                        sDescTributo = ""
                        bJuros = True
                        bMulta = True
                   End If
                  .Close
               End With
                               
                sNumProcesso = SubNull(!numprocesso)
                               
               If bJuros Then
                  nValorJuros = CalculaJuros(!ValorTributo + nValorCorrecao, Format(!DataVencimento, "dd/mm/yyyy"), mskData.Text)
               Else
                  nValorJuros = 0
               End If
               If bMulta Then
                  nValorMulta = CalculaMulta(!ValorTributo + nValorCorrecao, Format(!DataVencimento, "dd/mm/yyyy"), mskData.Text)
               Else
                  nValorMulta = 0
               End If
                If chkReparc.value = 1 Then nNumLivro = 0
               Sql = "INSERT DIVIDATIVA (USUARIO,NUMLIVRO,TIPOLIVRO,ANOLIVRO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,"
               Sql = Sql & "SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,DATAVENCIMENTO,PAGINALIVRO,DATAINSCRICAO,"
               Sql = Sql & "NUMCERTIDAO,CODTRIBUTO,VALORTRIBUTO,VALORJUROS,VALORMULTA,VALORCORRECAO,INSCRICAO,PROPRIETARIO,"
               Sql = Sql & "ENDERECO,COMPLEMENTO,CEP,BAIRRO,CIDADE,UF,DESCTRIBUTO,PROCESSO) VALUES('" & NomeDeLogin & "',"
               Sql = Sql & nNumLivro & "," & nCodTipo & "," & nAno & "," & !CODREDUZIDO & "," & !AnoExercicio & ","
               Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & ",'"
               Sql = Sql & Format(!DataVencimento, "mm/dd/yyyy") & "'," & Val(SubNull(!paginalivro)) & ",'" & Format(!datainscricao, "mm/dd/yyyy") & "',"
               Sql = Sql & Val(SubNull(!numcertidao)) & "," & !CodTributo & "," & Virg2Ponto(!ValorTributo) & ","
               Sql = Sql & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & "," & Virg2Ponto(CStr(nValorCorrecao)) & ",'" & sInscricao & "','"
               Sql = Sql & Mask(Left$(sProp, 40)) & "','" & Left$(Mask(sEndereco), 50) & "','" & Left$(sCompl, 35) & "','"
               Sql = Sql & sCep & "','" & Left(sBairro, 40) & "','" & Mask(sCidade) & "','" & sUF & "','" & Mask(sDescTributo) & "','" & sNumProcesso & "')"
               cn.Execute Sql, rdExecDirect
               nValorTotal = nValorTotal + !ValorTributo + nValorJuros + nValorJuros
               
                bFind = False
                For x = 0 To UBound(aSomaTributo)
                    If aSomaTributo(x).nNumLivro = nNumLivro And aSomaTributo(x).nCodTributo = !CodTributo Then
                        bFind = True
                        Exit For
                    End If
                Next
                
                If bFind Then
                    aSomaTributo(x).nSomaP = aSomaTributo(x).nSomaP + !ValorTributo
                    aSomaTributo(x).nSomaM = aSomaTributo(x).nSomaM + nValorMulta
                    aSomaTributo(x).nSomaJ = aSomaTributo(x).nSomaJ + nValorJuros
                    aSomaTributo(x).nSomaC = aSomaTributo(x).nSomaC + nValorCorrecao
                    aSomaTributo(x).nSomaT = aSomaTributo(x).nSomaT + !ValorTributo + nValorMulta + nValorJuros + nValorCorrecao
                Else
                    ReDim Preserve aSomaTributo(UBound(aSomaTributo) + 1)
                    aSomaTributo(UBound(aSomaTributo)).nNumLivro = nNumLivro
                    aSomaTributo(UBound(aSomaTributo)).nCodTributo = !CodTributo
                    aSomaTributo(UBound(aSomaTributo)).sDescTributo = sDescTributo
                    aSomaTributo(UBound(aSomaTributo)).nSomaP = !ValorTributo
                    aSomaTributo(UBound(aSomaTributo)).nSomaM = nValorMulta
                    aSomaTributo(UBound(aSomaTributo)).nSomaJ = nValorJuros
                    aSomaTributo(UBound(aSomaTributo)).nSomaC = nValorCorrecao
                    aSomaTributo(UBound(aSomaTributo)).nSomaT = !ValorTributo + nValorMulta + nValorJuros + nValorCorrecao
                End If
               
proximo:
               xId = xId + 1
               DoEvents
              .MoveNext
            Loop
           .Close
        End With
'        ReDim aSomaTributo(0)
        grdTemp.TextMatrix(z, 4) = FormatNumber(nValorTotal, 2)
        nValorTotal = 0
        CallPb CLng(z), .Rows - 1
PROXIMO2:
    Next
End With

For x = 0 To UBound(aSomaTributo)
    With aSomaTributo(x)
        If .nCodTributo > 0 Then
            Sql = "insert dividativatotal(USUARIO,numlivro,codigo,nome,somap,somam,somaj,somac,somat)values('" & NomeDeLogin & "',"
            Sql = Sql & .nNumLivro & "," & .nCodTributo & ",'" & .sDescTributo & "'," & Virg2Ponto(sTr(.nSomaP)) & "," & Virg2Ponto(sTr(.nSomaM)) & ","
            Sql = Sql & Virg2Ponto(sTr(.nSomaJ)) & "," & Virg2Ponto(sTr(.nSomaC)) & "," & Virg2Ponto(sTr(.nSomaT)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
    End With
Next

End Sub

Private Sub Calculo2()
Dim sStatus As String, RdoAux As rdoResultset, Sql As String, x As Integer, nAno As Integer, nCodTipo As Integer
Dim nNumLivro As Integer, xId As Long, nNumRec As Long, RdoEnd As rdoResultset, nCodReduz As Long, bFind As Boolean
Dim sNome As String, sEnd As String, sCompl As String, sCep As String, sBairro As String, sCidade As String, sUF As String
Dim qd As New rdoQuery, RdoDeb As rdoResultset, sCodLanc As String, bReparc As Boolean, aDebito() As Debito, nEval As Integer
Dim aTipo() As Long, sInscricao As String

ReDim aTipo(0)
For x = 0 To lstTipo.ListCount - 1
    If lstTipo.Selected(x) = True Then
        ReDim Preserve aTipo(UBound(aTipo) + 1)
        aTipo(UBound(aTipo)) = lstTipo.ItemData(lstTipo.ListIndex)
    End If
Next

bReparc = IIf(chkReparc.value = vbChecked, True, False)
sStatus = ""

For x = 0 To lstStatus.ListCount - 1
    If lstStatus.Selected(x) = True Then
        sStatus = sStatus & CStr(lstStatus.ItemData(x)) & ","
    End If
Next
If sStatus <> "" Then
    sStatus = Left(sStatus, Len(sStatus) - 1)
    sStatus = " AND STATUSLANC in (" & sStatus & ")"
End If

Sql = "delete  from DIVIDATIVA where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Sql = "delete  from DIVIDATIVATOTAL where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

For z = 1 To grdTemp.Rows - 1
    grdTemp.TopRow = z
    If bReparc Then
        If z > 1 Then
            nAnoLinha = grdTemp.TextMatrix(z, 0)
            If nAnoLinha = nAno Then
                GoTo NextBook
            End If
        End If
    End If

    nAno = grdTemp.TextMatrix(z, 0)
    nCodTipo = Val(Left$(grdTemp.TextMatrix(z, 1), 1))
    nNumLivro = grdTemp.TextMatrix(z, 3)
    xId = 0

    If nCodTipo = 1 Then
        If chkReparc.value = 1 Then
            Sql = "SELECT DISTINCT codreduzido From debitoparcela WHERE (codreduzido BETWEEN 1 AND 100000) AND ANOEXERCICIO=" & nAno & " AND "
            Sql = Sql & "codlancamento=20" & sStatus
        Else
            Sql = "SELECT DISTINCT codreduzido From debitoparcela WHERE (codreduzido BETWEEN 1 AND 100000) AND YEAR(datainscricao)=" & nAno & " AND "
            Sql = Sql & "codlancamento=1 AND numerolivro=" & nNumLivro & sStatus
        End If
    End If
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    Do Until RdoAux.EOF
         xId = RdoAux.AbsolutePosition
         If xId Mod 50 = 0 Then
            CallPb2 xId, RdoAux.RowCount
            DoEvents
         End If
         nCodReduz = RdoAux!CODREDUZIDO

         'Carrega o endereço
         If nCodTipo = 1 Then
            Sql = "select * from vwfullimovel where codreduzido=" & nCodReduz
            Set RdoEnd = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoEnd
                sInscricao = SubNull(!Inscricao)
                sNome = !NomeCidadao
                sEnd = !Logradouro & " Nº " & CStr(!Li_Num)
                sCompl = SubNull(Left(!Li_Compl, 35))
                sCep = RetornaCEP(!CodLogr, !Li_Num)
                sBairro = SubNull(!DescBairro)
                sCidade = "JABOTICABAL"
                sUF = "SP"
                .MoveNext
            End With
         End If

        'carrega a divida
        Set qd.ActiveConnection = cn
        qd.QueryTimeout = 0
        On Error Resume Next
        RdoDeb.Close
        On Error GoTo 0
        ReDim aDebito(0)

        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = nAno
        qd(3) = nAno
        qd(4) = 1
        qd(5) = 999
        qd(6) = 0
        qd(7) = 9999
        qd(8) = 1
        qd(9) = 999
        qd(10) = 0
        qd(11) = 999
        qd(12) = 1
        qd(13) = 999
        qd(14) = Format(Now, "mm/dd/yyyy")
        qd(15) = NomeDoUsuario
        qd(16) = 0
'        qd(17) = 1
        Set RdoDeb = qd.OpenResultset(rdOpenKeyset)
        With RdoDeb
            Do Until .EOF
                If bReparc And !CodLancamento <> 20 Then GoTo NextDebito
                If IsNull(!datainscricao) Then GoTo NextDebito
                If Not isInLongArray(aTipo(), CLng(!tributolivro)) Then GoTo NextDebito
                
                nEval = UBound(aDebito)
                bFind = False
                For x = 1 To nEval
                    If aDebito(x).nCodReduz = nCodReduz And aDebito(x).nAno = nAno And aDebito(x).nLanc = !CodLancamento And _
                       aDebito(x).nSeq = !SeqLancamento And _
                       aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                       bFind = True
                       Exit For
                    End If
                Next
                If Not bFind Then
                    ReDim Preserve aDebito(UBound(aDebito) + 1)
                    nEval = UBound(aDebito)
                    aDebito(nEval).nCodReduz = !CODREDUZIDO
                    aDebito(nEval).nAno = !AnoExercicio
                    aDebito(nEval).nLanc = !CodLancamento
                    aDebito(nEval).sLanc = !DESCLANCAMENTO
                    aDebito(nEval).nSeq = !SeqLancamento
                    aDebito(nEval).nParc = !NumParcela
                    aDebito(nEval).nCompl = !CODCOMPLEMENTO
                    aDebito(nEval).nSituacao = !statuslanc
                    aDebito(nEval).sSituacao = !Situacao
                    aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                    aDebito(nEval).nCodTributo = !CodTributo
                    aDebito(nEval).sDescTributo = SubNull(!ABREVTRIBUTO)
                    aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                    aDebito(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                    aDebito(nEval).nValorJuros = FormatNumber(!ValorJuros, 2)
                    aDebito(nEval).nValorMulta = FormatNumber(!ValorMulta, 2)
                    aDebito(nEval).nValorCorrecao = FormatNumber(!ValorCorrecao, 2)
                    aDebito(nEval).nPagina = !PAGINA
                    aDebito(nEval).nCertidao = !CERTIDAO
                    aDebito(nEval).dDataInscricao = Format(!datainscricao, "dd/mm/yyyy")
                    aDebito(nEval).sNumProcesso = SubNull(!numprocesso)
                Else
                    If aDebito(x).nCodTributo = !CodTributo Then GoTo NextDebito

                    aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
                    aDebito(x).nValorJuros = FormatNumber(aDebito(x).nValorJuros + !ValorJuros, 2)
                    aDebito(x).nValorMulta = FormatNumber(aDebito(x).nValorMulta + !ValorMulta, 2)
                    aDebito(x).nValorCorrecao = FormatNumber(aDebito(x).nValorCorrecao + !ValorCorrecao, 2)
                    aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTotal, 2)
                End If

NextDebito:
                xId = xId + 1
               .MoveNext

            Loop
           .Close
        End With
        
        For x = 1 To UBound(aDebito)
            With aDebito(x)
                Sql = "INSERT DIVIDATIVA (USUARIO,NUMLIVRO,TIPOLIVRO,ANOLIVRO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "DATAVENCIMENTO,PAGINALIVRO,DATAINSCRICAO,NUMCERTIDAO,CODTRIBUTO,VALORTRIBUTO,VALORJUROS,VALORMULTA,VALORCORRECAO,INSCRICAO,"
                Sql = Sql & "PROPRIETARIO,ENDERECO,COMPLEMENTO,CEP,BAIRRO,CIDADE,UF,DESCTRIBUTO,PROCESSO) VALUES('" & NomeDeLogin & "'," & nNumLivro & "," & nCodTipo & ","
                Sql = Sql & nAno & "," & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
                Sql = Sql & Format(.sVencto, "mm/dd/yyyy") & "'," & .nPagina & ",'" & Format(.dDataInscricao, "mm/dd/yyyy") & "'," & .nCertidao & ","
                Sql = Sql & .nCodTributo & "," & Virg2Ponto(CStr(.nValorTributo)) & "," & Virg2Ponto(CStr(.nValorJuros)) & "," & Virg2Ponto(CStr(.nValorMulta))
                Sql = Sql & "," & Virg2Ponto(CStr(.nValorCorrecao)) & ",'" & sInscricao & "','" & Mask(Left$(sNome, 40)) & "','" & Left$(Mask(sEnd), 50)
                Sql = Sql & "','" & Left$(sCompl, 35) & "','" & sCep & "','" & sBairro & "','" & sCidade & "','" & sUF & "','" & Mask(.sDescTributo) & "','" & sNumProcesso & "')"
                cn.Execute Sql, rdExecDirect
            End With
        Next

        RdoAux.MoveNext 'proximo codigo
    Loop
NextBook:
Next

    


End Sub


Private Sub CalculaCancel(eTipo As SeqEndereco)
Dim nAno As Integer, nCodTipo As Integer, nNumLivro As Integer, z As Integer, RdoAux As rdoResultset
Dim RdoC As rdoResultset, xId As Long, nNumRec As Long, nAnoLinha As Integer
Dim nValorCorrecao As Double, nValorTotal As Double, sDescTributo As String
Dim sProp As String, sStatus As String, sInscricao As String, sEndereco As String
Dim sCompl As String, sCep As String, sBairro As String, sCidade As String, sUF As String, sProcesso As String
Dim nValorJuros As Double, nValorMulta As Double, dDataInsc As Date, bJuros As Boolean, bMulta As Boolean
Dim aSomaTributo() As DIVIDATRIBUTO, bFind As Boolean
ReDim aSomaTributo(0)

DoEvents
Sql = "DELETE FROM DIVIDATIVA where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Sql = "delete  from DIVIDATIVATOTAL where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

nValorTotal = 0
xId = 0
        
Sql = "SELECT * FROM VWDIVIDAATIVAC "
If eTipo = Imobiliario Then
    Sql = Sql & "where codreduzido<100000"
ElseIf eTipo = Mobiliario Then
    Sql = Sql & "where codreduzido between 100000 and 500000"
Else
    Sql = Sql & "where codreduzido>500000"
End If
Set RdoC = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoC

    nNumRec = .RowCount
    If nNumRec = 0 Then Exit Sub
    
    If Not IsNull(!datainscricao) Then
        dDataInsc = Format(!datainscricao, "dd/mm/yyyy")
    Else
        dDataInsc = Format(Now, "dd/mm/yyyy")
    End If
    Do Until .EOF
       If xId Mod 10 = 0 Then
          CallPb2 xId, nNumRec
       End If
       sProcesso = SubNull(!NUMPROCESSO2)
       If !CODREDUZIDO < 100000 Then
          nCodTipo = 1
          Sql = "SELECT * FROM VWCNSIMOVEL WHERE CODREDUZIDO=" & !CODREDUZIDO
          Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          sEndereco = Trim$(SubNull(RdoAux!AbrevTipoLog)) & " " & Trim$(SubNull(RdoAux!AbrevTitLog)) & " " & RdoAux!NomeLogradouro & " nº " & SubNull(RdoAux!Li_Num)
          sCompl = SubNull(RdoAux!Li_Compl)
          sCep = RetornaCEP(Val(SubNull(RdoAux!CodLogr)), Val(SubNull(RdoAux!Li_Num)))
          If Trim(sCep) = "-" Then sCep = ""
          sBairro = SubNull(RdoAux!DescBairro)
          sCidade = SubNull(RdoAux!descCidade)
          sInscricao = RdoAux!Distrito & "." & Format(RdoAux!Setor, "00") & "." & Format(RdoAux!Quadra, "0000") & "." & Format(RdoAux!Lote, "00000") & "." & Format(RdoAux!Seq, "00") & "." & Format(RdoAux!Unidade, "00") & "." & Format(RdoAux!SubUnidade, "000")
          sUF = "SP"
          Sql = "SELECT NOMECIDADAO FROM VWCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & !CODREDUZIDO
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
          With RdoAux2
             If .RowCount = 0 Then
                sProp = ""
             Else
                sProp = !NomeCidadao
             End If
            .Close
          End With
       ElseIf !CODREDUZIDO >= 100000 And !CODREDUZIDO < 300000 Then
          nCodTipo = 2
          Sql = "SELECT vwLOGRADOURO.CODLOGRADOURO,vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG, vwLOGRADOURO.NOMELOGRADOURO, MOBILIARIO.NUMERO, "
          Sql = Sql & "MOBILIARIO.COMPLEMENTO, BAIRRO.DESCBAIRRO, CIDADE.DESCCIDADE, MOBILIARIO.SIGLAUF, MOBILIARIO.CEP,"
          Sql = Sql & "MOBILIARIO.NOMELOGRADOURO AS ENDANTIGO FROM MOBILIARIO INNER JOIN BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND "
          Sql = Sql & "MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO INNER JOIN CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
          Sql = Sql & "vwLOGRADOURO ON MOBILIARIO.CODLOGRADOURO = vwLOGRADOURO.CODLOGRADOURO WHERE CODIGOMOB=" & !CODREDUZIDO
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
          With RdoAux2
               If .RowCount > 0 Then
                    If Not IsNull(!NomeLogradouro) Then
                       sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " nº " & SubNull(!Numero)
                    Else
                       sEndereco = !ENDANTIGO & " nº " & SubNull(!Numero)
                    End If
                    sCompl = SubNull(!Complemento)
                    If SubNull(!descCidade) = "JABOTICABAL" Then
                       sCep = RetornaCEP(Val(SubNull(!CodLogradouro)), Val(SubNull(!Numero)))
                    Else
                       sCep = SubNull(!Cep)
                    End If
                    If Trim(sCep) = "-" Then sCep = ""
                    sBairro = !DescBairro
                    sCidade = !descCidade
                    sUF = !SiglaUF
               Else
                    sEndereco = ""
                    sCompl = ""
                    sCep = ""
                    sBairro = ""
                    sCidade = ""
                    sUF = ""
               End If
              .Close
          End With
          
          sInscricao = ""
          Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & !CODREDUZIDO
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
          With RdoAux2
             If .RowCount = 0 Then
                sProp = ""
             Else
                sProp = !RazaoSocial
             End If
            .Close
          End With
       ElseIf !CODREDUZIDO >= 500000 Then
          Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & !CODREDUZIDO
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
          With RdoAux2
             If .RowCount = 0 Then
                sProp = ""
             Else
                sProp = !NomeCidadao
             End If
            .Close
          End With
       End If
       
       If DateDiff("d", !DataVencimento, Now) > 0 Then
           nValorCorrecao = CalculaCorrecao(!ValorTributo, Format(!DataVencimento, "dd/mm/yyyy"))
       Else
           nValorCorrecao = 0
       End If
       
       Sql = "SELECT DESCTRIBUTO,MULTA,JUROS FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
           If .RowCount > 0 Then
                sDescTributo = !desctributo
                bJuros = !Juros
                bMulta = !Multa
           Else
                sDescTributo = ""
                bJuros = True
                bMulta = True
           End If
          .Close
       End With
                       
       If bJuros Then
          nValorJuros = CalculaJuros(!ValorTributo + nValorCorrecao, Format(!DataVencimento, "dd/mm/yyyy"))
       Else
          nValorJuros = 0
       End If
       If bMulta Then
          nValorMulta = CalculaMulta(!ValorTributo + nValorCorrecao, Format(!DataVencimento, "dd/mm/yyyy"))
       Else
          nValorMulta = 0
       End If
       Sql = "INSERT DIVIDATIVA (usuario,NUMLIVRO,TIPOLIVRO,ANOLIVRO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,"
       Sql = Sql & "SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,DATAVENCIMENTO,PAGINALIVRO,DATAINSCRICAO,"
       Sql = Sql & "NUMCERTIDAO,CODTRIBUTO,VALORTRIBUTO,VALORJUROS,VALORMULTA,VALORCORRECAO,INSCRICAO,PROPRIETARIO,"
       Sql = Sql & "ENDERECO,COMPLEMENTO,CEP,BAIRRO,CIDADE,UF,DESCTRIBUTO,PROCESSO) VALUES('" & NomeDeLogin & "',"
       Sql = Sql & Val(SubNull(!numerolivro)) & "," & nCodTipo & "," & 2007 & "," & !CODREDUZIDO & "," & !AnoExercicio & ","
       Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & ",'"
       Sql = Sql & Format(!DataVencimento, "mm/dd/yyyy") & "'," & Val(SubNull(!paginalivro)) & ",'" & Format(!datainscricao, "mm/dd/yyyy") & "',"
       Sql = Sql & Val(SubNull(!numcertidao)) & "," & !CodTributo & "," & Virg2Ponto(!ValorTributo) & ","
       Sql = Sql & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & "," & Virg2Ponto(CStr(nValorCorrecao)) & ",'" & sInscricao & "','"
       Sql = Sql & Mask(Left$(sProp, 40)) & "','" & Left$(Mask(sEndereco), 50) & "','" & Left$(sCompl, 35) & "','"
       Sql = Sql & sCep & "','" & sBairro & "','" & sCidade & "','" & sUF & "','" & sDescTributo & "','" & sProcesso & "')"
       cn.Execute Sql, rdExecDirect
       nValorTotal = nValorTotal + !ValorTributo + nValorJuros + nValorJuros
proximo:
       
                If Val(SubNull(!numerolivro)) > 10000 Then
                    nNumLivro = 0
                Else
                    nNumLivro = Val(SubNull(!numerolivro))
                End If
                bFind = False
                For x = 0 To UBound(aSomaTributo)
                    If aSomaTributo(x).nNumLivro = nNumLivro And aSomaTributo(x).nCodTributo = !CodTributo Then
                        bFind = True
                        Exit For
                    End If
                Next
                
                If bFind Then
                    aSomaTributo(x).nSomaP = aSomaTributo(x).nSomaP + !ValorTributo
                    aSomaTributo(x).nSomaM = aSomaTributo(x).nSomaM + nValorMulta
                    aSomaTributo(x).nSomaJ = aSomaTributo(x).nSomaJ + nValorJuros
                    aSomaTributo(x).nSomaC = aSomaTributo(x).nSomaC + nValorCorrecao
                    aSomaTributo(x).nSomaT = aSomaTributo(x).nSomaT + !ValorTributo + nValorMulta + nValorJuros + nValorCorrecao
                Else
                    ReDim Preserve aSomaTributo(UBound(aSomaTributo) + 1)
                    aSomaTributo(UBound(aSomaTributo)).nNumLivro = nNumLivro
                    aSomaTributo(UBound(aSomaTributo)).nCodTributo = !CodTributo
                    aSomaTributo(UBound(aSomaTributo)).sDescTributo = sDescTributo
                    aSomaTributo(UBound(aSomaTributo)).nSomaP = !ValorTributo
                    aSomaTributo(UBound(aSomaTributo)).nSomaM = nValorMulta
                    aSomaTributo(UBound(aSomaTributo)).nSomaJ = nValorJuros
                    aSomaTributo(UBound(aSomaTributo)).nSomaC = nValorCorrecao
                    aSomaTributo(UBound(aSomaTributo)).nSomaT = !ValorTributo + nValorMulta + nValorJuros + nValorCorrecao
                End If
       
       
       
       xId = xId + 1
      .MoveNext
    Loop
   .Close
End With

For x = 0 To UBound(aSomaTributo)
    With aSomaTributo(x)
        If .nCodTributo > 0 Then
            Sql = "insert dividativatotal(usuario,numlivro,codigo,nome,somap,somam,somaj,somac,somat)values('" & NomeDeLogin & "',"
            Sql = Sql & .nNumLivro & "," & .nCodTributo & ",'" & .sDescTributo & "'," & Virg2Ponto(sTr(.nSomaP)) & "," & Virg2Ponto(sTr(.nSomaM)) & ","
            Sql = Sql & Virg2Ponto(sTr(.nSomaJ)) & "," & Virg2Ponto(sTr(.nSomaC)) & "," & Virg2Ponto(sTr(.nSomaT)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
    End With
Next


nValorTotal = 0
frmReport.ShowReport "DIVIDATIVACANCELADO", frmMdi.HWND, Me.HWND

Sql = "DELETE FROM DIVIDATIVA where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Sql = "delete  from DIVIDATIVATOTAL where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2) & " %"

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub CallPb2(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If nTotal = 0 Then Exit Sub
If ((nPosF * 100) / nTotal) <= 100 Then
   Pb2.value = (nPosF * 100) / nTotal
Else
   Pb2.value = 100
End If
lblPB2.Caption = FormatNumber(Pb2.value, 2) & " %"

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub mnuCidadao_Click()
CalculaCancel cidadao
End Sub

Private Sub mnuEmpresa_Click()
CalculaCancel Mobiliario
End Sub

Private Sub mnuImovel_Click()
CalculaCancel Imobiliario
End Sub

Private Sub mskData_GotFocus()
mskData.SetFocus
End Sub
