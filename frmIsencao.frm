VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmIsencao 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Imunidade/Isenção"
   ClientHeight    =   3540
   ClientLeft      =   2385
   ClientTop       =   3555
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   8910
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Vitalício"
      Height          =   210
      Index           =   0
      Left            =   5760
      TabIndex        =   23
      Top             =   1050
      Width           =   870
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Anual"
      Height          =   210
      Index           =   1
      Left            =   6810
      TabIndex        =   22
      Top             =   1080
      Value           =   -1  'True
      Width           =   810
   End
   Begin VB.CheckBox chkFil 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Filantrópico"
      Height          =   210
      Left            =   4380
      TabIndex        =   21
      Top             =   1110
      Width           =   1200
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmIsencao.frx":0000
      Left            =   4965
      List            =   "frmIsencao.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   330
      Width           =   2565
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   0
      Top             =   300
      Width           =   1305
   End
   Begin prjChameleon.chameleonButton cmdAdd 
      Height          =   300
      Left            =   2985
      TabIndex        =   14
      ToolTipText     =   "Adicionar na Lista"
      Top             =   1845
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      BTYPE           =   6
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmIsencao.frx":0004
      PICN            =   "frmIsencao.frx":0020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDel 
      Height          =   300
      Left            =   2970
      TabIndex        =   15
      ToolTipText     =   "Remover da Lista"
      Top             =   2205
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      BTYPE           =   6
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmIsencao.frx":017A
      PICN            =   "frmIsencao.frx":0196
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
      Height          =   1320
      Left            =   3570
      TabIndex        =   16
      Top             =   1635
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   2328
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      BackColorFixed  =   15658734
      BackColorBkg    =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmIsencao.frx":02F0
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   7695
      TabIndex        =   11
      ToolTipText     =   "Sair da Tela"
      Top             =   3150
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
      MICON           =   "frmIsencao.frx":0382
      PICN            =   "frmIsencao.frx":039E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   6630
      TabIndex        =   12
      ToolTipText     =   "Gravar os Dados"
      Top             =   3150
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmIsencao.frx":040C
      PICN            =   "frmIsencao.frx":0428
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPerc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   4
      Top             =   2055
      Width           =   1305
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1365
      Width           =   1305
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1020
      Width           =   885
   End
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   90
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2610
      Width           =   2685
   End
   Begin esMaskEdit.esMaskedEdit mskDataProc 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1710
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      BackColor       =   15658734
      MouseIcon       =   "frmIsencao.frx":07CD
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
      Locked          =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdEtiqueta 
      Height          =   345
      Left            =   3570
      TabIndex        =   24
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   3120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Etiquetas"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIsencao.frx":07E9
      PICN            =   "frmIsencao.frx":0805
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduz.:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   345
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Isenção.:"
      Height          =   195
      Index           =   2
      Left            =   3630
      TabIndex        =   19
      Top             =   375
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome/Razão..:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   675
      Width           =   1140
   End
   Begin VB.Label lblNome 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1290
      TabIndex        =   17
      Top             =   675
      Width           =   5340
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano de Isenção..:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1065
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "% de Isenção.....:"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   9
      Top             =   2100
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Processo...:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1755
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo..:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1410
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo..:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   645
   End
End
Attribute VB_Name = "frmIsencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String
Dim xImovel As clsImovel, bIPTU As Boolean
Dim sRet As String
Dim evEdit As Integer
Dim bEdit As Boolean

Private Sub chkFil_LostFocus()
Opt(0).SetFocus
End Sub

Private Sub cmbTipo_Click()
If Val(txtCod.Text) = 0 Then Exit Sub
If cmbTipo.ListIndex = -1 Then Exit Sub
If bIPTU Then
    Select Case cmbTipo.ItemData(cmbTipo.ListIndex)
        Case 1
            Opt(0).value = True
            txtAno.SetFocus
        Case 4
            MsgBox "Um imóvel não pode ser isento de ISS.", vbCritical, "Atenção"
            cmbTipo.ListIndex = -1
        Case 5
            MsgBox "Um imóvel não pode ser isento de Taxa de Licença.", vbCritical, "Atenção"
            cmbTipo.ListIndex = -1
    End Select
Else
    Select Case cmbTipo.ItemData(cmbTipo.ListIndex)
        Case 1
            Opt(0).value = True
            txtAno.SetFocus
        Case 3
            MsgBox "Uma empresa não pode ser isenta de IPTU.", vbCritical, "Atenção"
            cmbTipo.ListIndex = -1
    End Select
End If

End Sub

Private Sub cmdAdd_Click()
Dim sTipo As String, sPeriodo As String

If Val(txtCod.Text) = 0 Then
    MsgBox "Digite o código do contribuinte", vbExclamation, "Atenção"
    Exit Sub
End If

If cmbTipo.ListIndex = -1 Then
    MsgBox "Selecione o Tipo de Isenção", vbExclamation, "Atenção"
    Exit Sub
End If

Select Case cmbTipo.ItemData(cmbTipo.ListIndex)
    Case 1
        sTipo = "IMUN"
    Case 3
        sTipo = "IPTU"
    Case 4
        sTipo = "ISS"
    Case 5
        sTipo = "TX.LIC"
End Select

If Opt(0).value = True Then
   sPeriodo = "Vit"
ElseIf Opt(1).value = True Then
   sPeriodo = "Anu"
End If

If sTipo = "IMUN" And Opt(0).value = False Then
    MsgBox "Imunidade requer período vitalício", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno.Text) = 0 Then
    MsgBox "Digite o Ano de Isenção", vbExtender, "Atenção"
    Exit Sub
End If

If Trim(txtNumProc.Text) = "" And cmbTipo.ListIndex > 0 Then
    MsgBox "Digite o Nº do Processo", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataProc.Text) And cmbTipo.ListIndex > 0 Then
    MsgBox "Data do Processo Inválida", vbExclamation, "Atenção"
    Exit Sub
End If

If sTipo = "IMUN" Then
    txtPerc.Text = 100
Else
    If Val(txtPerc.Text) < 1 Or Val(txtPerc.Text) > 100 Then
        MsgBox "Digite o % de Isenção entre 1 e 100", vbExclamation, "Atenção"
        Exit Sub
    End If
End If
If grdTemp.Rows > 1 Then
    'If Val(txtAno.text) < grdTemp.TextMatrix(grdTemp.Rows - 1, 0) Then
    '   MsgBox "O ano de isenção tem que ser maior ou igual a última isenção atribuida.", vbCritical, "Atenção"
    '   Exit Sub
    'End If
    If Val(txtAno.Text) = grdTemp.TextMatrix(grdTemp.Rows - 1, 0) And sTipo = grdTemp.TextMatrix(grdTemp.Rows - 1, 1) Then
       MsgBox "Este tipo de isenção já foi atribuida para o ano especificado.", vbCritical, "Atenção"
       Exit Sub
    End If
End If

If Opt(0).value = False And Opt(1).value = False Then
   MsgBox "Selecione o período.", vbCritical, "Atenção"
   Exit Sub
End If

If sTipo = "IMUN" Then
   grdTemp.AddItem txtAno.Text & Chr(9) & sTipo & Chr(9) & sPeriodo & Chr(9) & chkFil.value & Chr(9) & 100 & Chr(9) & Mask(txtNumProc.Text) & Chr(9) & mskDataProc.Text & Chr(9) & Mask(txtMotivo.Text) & Chr(9) & NomeDeLogin & Chr(9) & Format(Now, "dd/mm/yyyy")
Else
   grdTemp.AddItem txtAno.Text & Chr(9) & sTipo & Chr(9) & sPeriodo & Chr(9) & chkFil.value & Chr(9) & txtPerc.Text & Chr(9) & Mask(txtNumProc.Text) & Chr(9) & mskDataProc.Text & Chr(9) & Mask(txtMotivo.Text) & Chr(9) & NomeDeLogin & Chr(9) & Format(Now, "dd/mm/yyyy")
End If

txtAno.Text = ""
txtNumProc.Text = ""
LimpaMascara mskDataProc
txtPerc.Text = ""
txtMotivo.Text = ""

cmdGravar.Enabled = True
End Sub

Private Sub cmdDel_Click()
Dim nAno As Integer, nSeq As Integer, s As String

If grdTemp.Rows = 1 Then Exit Sub

nAno = Val(grdTemp.TextMatrix(grdTemp.Row, 0))

If MsgBox("Excluir esta Isenção ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    If grdTemp.Rows > 2 Then
       grdTemp.RemoveItem grdTemp.Row
    Else
       grdTemp.Rows = 1
    End If
    cmdGravar.Enabled = True
End If

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With

s = "Excluída isenção do imóvel para o ano de " & nAno & " por " & NomeDeLogin

'Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & Mask(s) & "','GTI','" & Format(Now, "mm/dd/yyyy") & "')"
Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & Mask(s) & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
cn.Execute Sql, rdExecDirect
cmdGravar_Click

End Sub

Private Sub cmdEtiqueta_Click()

'variaveis para arquivo texto
Dim sExercicio As String, sContribuinte As String, sSacado As String, sEnd As String, sCompl As String, sBairro As String
Dim sQuadra As String, sCep As String, sLote As String, sEndEntrega As String, sComplEntrega As String, sBairroEntrega As String
Dim sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sCodContribuinte As String, sInscricao As String
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, xId As Integer, sCodInscricao As String

Ocupado
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT processogti.CODCIDADAO,NUMERO FROM processogti INNER JOIN assunto ON processogti.CODASSUNTO = assunto.CODIGO LEFT OUTER JOIN "
Sql = Sql & "centrocusto ON processogti.CENTROCUSTO = centrocusto.CODIGO LEFT OUTER JOIN cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
Sql = Sql & "WHERE (processogti.ANO BETWEEN 2008 AND 2008) AND (processogti.CODASSUNTO = 759 OR processogti.CODASSUNTO = 828) "
Sql = Sql & "ORDER BY cidadao.codcidadao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
xId = 1
Do Until .EOF
    Sql = "SELECT codreduzido, anoisencao, codisencao, numprocesso, dataprocesso, percisencao, filantropico, periodo, motivo From isencao "
    Sql = Sql & "WHERE (anoisencao = 2009) AND (numprocesso = '" & !Numero & "/2008')"
    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux3.RowCount = 0 Then GoTo Proximo
    
    
    Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & !CodCidadao
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        sCodInscricao = !CodCidadao
        sContribuinte = !NomeCidadao
        sEndEntrega = !Endereco & ", " & !NUMIMOVEL
        sComplEntrega = !Complemento
        If !CodLogradouro > 0 Then
            sCepEntrega = RetornaCEP(!CodLogradouro, !NUMIMOVEL)
        Else
            sCepEntrega = "00000-000"
        End If
        sBairroEntrega = SubNull(!DescBairro)
        sCidEntrega = SubNull(!descCidade)
        sUFEntrega = SubNull(!SiglaUF)
       .Close
    End With
    
    xId = xId + 1
    Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
    Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & Mask(sContribuinte) & "','"
    Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
    cn.Execute Sql, rdExecDirect
    xId = xId + 1
Proximo:
   .MoveNext
    Loop
   .Close
End With
Liberado
frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, nCodIsencao As Integer, nCodPeriodo As Integer, nAnoproc As Integer, nNumproc As Long
Sql = "DELETE FROM ISENCAO WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

With grdTemp
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 1) = "IMUN" Then
            nCodIsencao = 1
        ElseIf .TextMatrix(x, 1) = "IPTU" Then
            nCodIsencao = 3
        ElseIf .TextMatrix(x, 1) = "ISS" Then
            nCodIsencao = 4
        ElseIf .TextMatrix(x, 1) = "TX.LIC" Then
            nCodIsencao = 5
        End If
        If .TextMatrix(x, 2) = "Vit" Then
           nCodPeriodo = 0
        ElseIf .TextMatrix(x, 2) = "Anu" Then
           nCodPeriodo = 1
        End If
        If Not IsNumeric(.TextMatrix(x, 6)) Then
            .TextMatrix(x, 6) = Format(Now, "dd/mm/yyyy")
        End If
        If .TextMatrix(x, 5) <> "" Then
            nNumproc = Val(Left$(.TextMatrix(x, 5), Len(.TextMatrix(x, 5)) - 2))
            nAnoproc = Right(.TextMatrix(x, 5), 4)
        Else
            nNumproc = 0
            nAnoproc = 0
        End If
'        Sql = "INSERT ISENCAO (CODREDUZIDO,ANOISENCAO,CODISENCAO,NUMPROCESSO,PERCISENCAO,FILANTROPICO,PERIODO,MOTIVO,ANOPROC,NUMPROC,USUARIO,DATAALTERA) VALUES("
'        Sql = Sql & Val(txtCod.Text) & "," & Val(.TextMatrix(x, 0)) & "," & nCodIsencao & ",'" & .TextMatrix(x, 5) & "',"
'        Sql = Sql & Virg2Ponto(.TextMatrix(x, 4)) & "," & Val(.TextMatrix(x, 3)) & "," & nCodPeriodo & ",'" & .TextMatrix(x, 7) & "'," & nAnoproc & "," & nNumproc & ",'"
'        Sql = Sql & NomeDeLogin & "','" & Format(Now, "mm/dd/yyyy") & "')"
        Sql = "INSERT ISENCAO (CODREDUZIDO,ANOISENCAO,CODISENCAO,NUMPROCESSO,PERCISENCAO,FILANTROPICO,PERIODO,MOTIVO,ANOPROC,NUMPROC,USERID,DATAALTERA) VALUES("
        Sql = Sql & Val(txtCod.Text) & "," & Val(.TextMatrix(x, 0)) & "," & nCodIsencao & ",'" & .TextMatrix(x, 5) & "',"
        Sql = Sql & Virg2Ponto(.TextMatrix(x, 4)) & "," & Val(.TextMatrix(x, 3)) & "," & nCodPeriodo & ",'" & .TextMatrix(x, 7) & "'," & nAnoproc & "," & nNumproc & ","
        Sql = Sql & RetornaUsuarioID(.TextMatrix(x, 8)) & ",'" & Format(.TextMatrix(x, 9), "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    Next
End With

cmdGravar.Enabled = False
End Sub


Private Sub cmdSair_Click()

If cmdGravar.Enabled = True Then
    If MsgBox("Salvar as alterações ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        cmdGravar_Click
    Else
        cmdGravar.Enabled = False
    End If
End If

Unload Me
End Sub

Private Function FileDialog1_OnSelectionChanged(ByVal SelectedPageID As Long, pvSelectedObject As Variant) As Boolean

End Function

Private Sub Form_Load()
Centraliza Me
Ocupado
sRet = RetEventUserForm(Me.Name)
FormHagana

Sql = "SELECT CODTIPO,DESCTIPO FROM TIPOISENCAO WHERE CODTIPO<>0 AND CODTIPO<>2"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !DESCTIPO
        cmbTipo.ItemData(cmbTipo.NewIndex) = !CodTipo
       .MoveNext
    Loop
   .Close
End With
cmbTipo.ListIndex = 1
Set xImovel = New clsImovel
'cmdGravar.Enabled = False
'txtCod.SetFocus
Liberado
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub Limpa()
txtAno.Text = ""
lblNome.Caption = ""
cmbTipo.ListIndex = 1
txtNumProc.Text = ""
LimpaMascara mskDataProc
txtPerc.Text = ""
txtMotivo.Text = ""
grdTemp.Rows = 1
Opt(0).value = False
Opt(1).value = False
chkFil.value = 0
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

'If cmdGravar.Enabled = True Then
'    If MsgBox("Salvar as alterações ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'        KeyAscii = 0
'        cmdGravar_Click
'    Else
'        cmdGravar.Enabled = False
'    End If
'End If

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    txtCod_LostFocus
Else
    Tweak txtCod, KeyAscii, IntegerPositive
    
End If
End Sub
Private Sub txtCod_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtCod.Text) = 0 Then Limpa
End Sub

Private Sub txtCod_LostFocus()
    If Val(txtCod.Text) > 0 Then
    CarregaEmpresa
    If lblNome.Caption <> "" Then
       CarregaIsencao
    End If
    cmdGravar.Enabled = False
End If
txtAno.Text = Year(Now) + 1
Opt(1).value = True
txtPerc.Text = 100
txtNumProc.SetFocus
End Sub

Private Sub CarregaIsencao()
Dim sTipo As String
Dim sPeriodo As String
'693
Sql = "SELECT * FROM VWISENCAOPROCESSO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " ORDER BY ANOISENCAO DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Do Until .EOF
            If !CODISENCAO = 1 Then
                sTipo = "IMUN"
            ElseIf !CODISENCAO = 3 Then
                sTipo = "IPTU"
            ElseIf !CODISENCAO = 4 Then
                sTipo = "ISS"
            ElseIf !CODISENCAO = 5 Then
                sTipo = "TX.LIC"
            Else
                sTipo = ""
            End If
           
            If Val(SubNull(!PERIODO)) = 0 Then
               sPeriodo = "Vit"
            ElseIf Val(SubNull(!PERIODO)) = 1 Then
               sPeriodo = "Anu"
            End If
            
            If IsNull(!dataaltera) Then
                grdTemp.AddItem !anoisencao & Chr(9) & sTipo & Chr(9) & sPeriodo & Chr(9) & Val(SubNull(!FILANTROPICO)) & Chr(9) & !percisencao & Chr(9) & SubNull(!numprocesso) & Chr(9) & IIf(IsNull(!DATAPROCESSO), "", Format(!DATAPROCESSO, "dd/mm/yyyy")) & Chr(9) & SubNull(!MOTIVO) & Chr(9) & SubNull(!NomeLogin)
            Else
                grdTemp.AddItem !anoisencao & Chr(9) & sTipo & Chr(9) & sPeriodo & Chr(9) & Val(SubNull(!FILANTROPICO)) & Chr(9) & !percisencao & Chr(9) & SubNull(!numprocesso) & Chr(9) & IIf(IsNull(!DATAPROCESSO), "", Format(!DATAPROCESSO, "dd/mm/yyyy")) & Chr(9) & SubNull(!MOTIVO) & Chr(9) & SubNull(!NomeLogin) & Chr(9) & Format(!dataaltera, "dd/mm/yyyy")
            End If
           .MoveNext
        Loop
    End If
End With

End Sub

Private Sub CarregaEmpresa()
Dim nCodReduz As Long

nCodReduz = Val(txtCod.Text)

Limpa
Sql = "SELECT CODIGOMOB,RAZAOSOCIAL,DATAENCERRAMENTO "
Sql = Sql & " FROM MOBILIARIO WHERE CODIGOMOB=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
       If Not IsNull(!dataencerramento) Or !dataencerramento <> CDate("01/01/1900") Then
          MsgBox "Esta empresa foi encerrada em " & Format(!dataencerramento, "dd/mm/yyyy"), vbExclamation, "Atenção"
          Exit Sub
       End If
      'suspenção
       Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & txtCod.Text
       Sql = Sql & " ORDER BY DATAEVENTO DESC"
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
           If .RowCount > 0 Then
               If !CODTIPOEVENTO = 2 Then
                   MsgBox "Esta empresa esta SUSPENSA", vbExclamation, "Atenção"
                   Exit Sub
               End If
           End If
          .Close
       End With
       lblNome.Caption = !RazaoSocial
       bIPTU = False
       cmbTipo.SetFocus
    Else
       With xImovel
            .CarregaImovel nCodReduz
            If .CodigoImovel > 0 Then
                bIPTU = True
                lblNome.Caption = .NomePropPrincipal
                cmbTipo.SetFocus
            Else
                MsgBox "Não existe Imóvel nem Empresa com este Código.", vbExclamation, "Atenção"
'                txtCod.SetFocus
            End If
       End With
    End If
End With

End Sub

Private Sub txtNumProc_LostFocus()
Dim sValidaProc As String

sValidaProc = ValidaProcesso(txtNumProc.Text)

If txtNumProc.Text <> "" Then
    'If sValidaProc = "OK" Then
    If InStr(1, sValidaProc, "ARQUIVADO", vbBinaryCompare) = 0 And InStr(1, sValidaProc, "CANCELADO", vbBinaryCompare) = 0 And sValidaProc <> "OK" Then
        MsgBox sValidaProc, vbExclamation, "Atenção"
        LimpaMascara mskDataProc
    Else
        mskDataProc.Text = Format(RetornaDataProcesso(Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 5)), Val(Right$(txtNumProc.Text, 4))), "dd/mm/yyyy")
    End If
Else
    LimpaMascara mskDataProc
End If

End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
Tweak txtPerc, KeyAscii, DecimalPositive
End Sub

Private Sub FormHagana()

evEdit = 3

If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True

If Not bEdit Then cmdAdd.Enabled = False
If Not bEdit Then cmdDel.Enabled = False
If Not bEdit Then cmdGravar.Enabled = False
If bEdit Then cmdGravar.Enabled = True
End Sub

