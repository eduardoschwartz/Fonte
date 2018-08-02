VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmResumoProtocolo 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo de Processos Cadastrados"
   ClientHeight    =   4800
   ClientLeft      =   7620
   ClientTop       =   3150
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   6045
   Begin VB.ListBox lstNomeLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   1590
      ItemData        =   "frmResumoProtocolo.frx":0000
      Left            =   1170
      List            =   "frmResumoProtocolo.frx":0002
      TabIndex        =   6
      Top             =   2385
      Visible         =   0   'False
      Width           =   4515
   End
   Begin VB.TextBox txtNomeLogr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1170
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2385
      Width           =   3720
   End
   Begin VB.TextBox txtCodLogr 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   4950
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2385
      Width           =   735
   End
   Begin VB.Frame frDDList 
      BackColor       =   &H00EEEEEE&
      Height          =   375
      Left            =   1125
      TabIndex        =   21
      Top             =   1395
      Width           =   4560
      Begin prjChameleon.chameleonButton cmdDDList 
         Height          =   240
         Left            =   1845
         TabIndex        =   22
         ToolTipText     =   "Exibir Lista"
         Top             =   45
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         BTYPE           =   14
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   14869218
         BCOLO           =   14869218
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmResumoProtocolo.frx":0004
         PICN            =   "frmResumoProtocolo.frx":0020
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvDDList 
         Height          =   2100
         Left            =   45
         TabIndex        =   23
         Top             =   405
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   3704
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Ativo"
            Object.Width           =   953
         EndProperty
      End
   End
   Begin VB.ComboBox cmbArquivado 
      Height          =   315
      ItemData        =   "frmResumoProtocolo.frx":017A
      Left            =   1170
      List            =   "frmResumoProtocolo.frx":0187
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3435
      Width           =   1245
   End
   Begin VB.CheckBox chkExterno 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Somente Externos"
      Height          =   195
      Left            =   2310
      TabIndex        =   9
      Top             =   2865
      Width           =   1665
   End
   Begin VB.ComboBox cmbSetor 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1890
      Width           =   4605
   End
   Begin VB.TextBox txtAssunto 
      Height          =   315
      Left            =   3870
      TabIndex        =   16
      Top             =   750
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CheckBox chkTramite 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Último Trâmite"
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   3915
      Width           =   1560
   End
   Begin VB.ComboBox cmbOrder 
      Height          =   315
      ItemData        =   "frmResumoProtocolo.frx":019E
      Left            =   1140
      List            =   "frmResumoProtocolo.frx":01AE
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Width           =   2265
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Imprimir Todos os Assuntos"
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   675
      Value           =   1  'Checked
      Width           =   2310
   End
   Begin esMaskEdit.esMaskedEdit mskDataDe 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmResumoProtocolo.frx":01E6
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
   Begin esMaskEdit.esMaskedEdit mskDataAte 
      Height          =   285
      Left            =   3765
      TabIndex        =   1
      Top             =   240
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmResumoProtocolo.frx":0202
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
   Begin prjChameleon.chameleonButton cmdReq 
      Height          =   345
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Sair da Tela"
      Top             =   2805
      Width           =   465
      _ExtentX        =   820
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmResumoProtocolo.frx":021E
      PICN            =   "frmResumoProtocolo.frx":023A
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
      Cancel          =   -1  'True
      Height          =   345
      Left            =   1710
      TabIndex        =   8
      ToolTipText     =   "Sair da Tela"
      Top             =   2805
      Width           =   465
      _ExtentX        =   820
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmResumoProtocolo.frx":0394
      PICN            =   "frmResumoProtocolo.frx":03B0
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
      Left            =   4530
      TabIndex        =   26
      Top             =   4365
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmResumoProtocolo.frx":050A
      PICN            =   "frmResumoProtocolo.frx":0526
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3285
      TabIndex        =   27
      Top             =   4365
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmResumoProtocolo.frx":0594
      PICN            =   "frmResumoProtocolo.frx":05B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEtiqueta 
      Height          =   345
      Left            =   90
      TabIndex        =   28
      ToolTipText     =   "Imprimir Mala Direta"
      Top             =   4365
      Width           =   1185
      _ExtentX        =   2090
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmResumoProtocolo.frx":070A
      PICN            =   "frmResumoProtocolo.frx":0726
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
      Caption         =   "Logradouro...:"
      Height          =   225
      Index           =   1
      Left            =   90
      TabIndex        =   25
      Top             =   2415
      Width           =   1050
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivados.:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   3495
      Width           =   1080
   End
   Begin VB.Label lblReq 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   225
      Left            =   1290
      TabIndex        =   19
      Top             =   3195
      Width           =   4485
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Requerente.:"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   2895
      Width           =   1020
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Setor........:"
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   17
      Top             =   1980
      Width           =   840
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenar por:"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   15
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Período até:"
      Height          =   195
      Index           =   7
      Left            =   2760
      TabIndex        =   14
      Top             =   300
      Width           =   930
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Período de:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   13
      Top             =   300
      Width           =   840
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto....:"
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   12
      Top             =   1530
      Width           =   840
   End
End
Attribute VB_Name = "frmResumoProtocolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset

Private Sub chkAll_Click()
If chkAll.Value = vbChecked Then
    CheckAll
    cmdDDList.Enabled = False
Else
    CheckNone
    cmdDDList.Enabled = True
End If
End Sub


Private Sub cmdDDList_Click()
If cmdDDList.Value = True Then
    frDDList.ZOrder 0
    frDDList.Height = 2580
Else
    frDDList.Height = 375
End If
End Sub

Private Sub cmdDel_Click()
lblReq.Caption = "0"
End Sub

Private Sub cmdEtiqueta_Click()
Dim x As Integer, sAssunto As String, aCodigo() As Long
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim xId As Long, nNumRec As Long, nCodLogr As Long, sCodInscricao As String, sContribuinte As String
Dim sEnd As String, nNum As Integer, sCEP As String, sCompl As String, sBairro As String
Dim sEndEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String

ReDim aCodigo(0)

If IsDate(mskDataDe.Text) And Not IsDate(mskDataAte.Text) Then
   MsgBox "Data inicial sem data final.", vbExclamation, "Atenção"
    Exit Sub
End If

sAssunto = "": txtAssunto.Text = ""
For x = 1 To lvDDList.ListItems.Count
    If lvDDList.ListItems(x).Checked = True Then
        sAssunto = sAssunto & CStr(Val(lvDDList.ListItems(x).SubItems(1))) & ","
    End If
Next

If sAssunto = "" Then
    MsgBox "Selecione um Assunto.", vbExclamation, "Atenção"
    Exit Sub
End If
sAssunto = Left$(sAssunto, Len(sAssunto) - 1)
txtAssunto.Text = sAssunto
If chkTramite.Value = 1 Then
    frmReport.ShowReport "TRAMITEABERTOLOCAL", frmMdi.hwnd, Me.hwnd
    Exit Sub
End If

If Not IsDate(mskDataDe.Text) Then
   MsgBox "Data inicial inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataAte.Text) Then
   MsgBox "Data Final inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If CDate(mskDataDe.Text) > CDate(mskDataAte.Text) Then
   MsgBox "Data inicial maior que a final.", vbExclamation, "Atenção"
   Exit Sub
End If
             
Sql = "SELECT processogti.CodCidadao FROM processogti INNER JOIN assunto ON processogti.CODASSUNTO = assunto.CODIGO LEFT OUTER JOIN vwPROCESSOENDERECO ON "
Sql = Sql & "processogti.ANO = vwPROCESSOENDERECO.ANO AND processogti.NUMERO = vwPROCESSOENDERECO.NUMPROCESSO LEFT OUTER JOIN centrocusto ON "
Sql = Sql & "processogti.CENTROCUSTO = centrocusto.CODIGO LEFT OUTER JOIN cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
Sql = Sql & "WHERE DATAENTRADA BETWEEN '" & Format(mskDataDe.Text, "mm/dd/yyyy") & "' AND '" & Format(mskDataAte.Text, "mm/dd/yyyy") & "' AND CODASSUNTO in (" & sAssunto & ")"
If cmbArquivado.ListIndex = 1 Then
    Sql = Sql & " AND DATAARQUIVA IS NOT NULL"
ElseIf cmbArquivado.ListIndex = 2 Then
    Sql = Sql & "  AND DATAARQUIVA IS NULL"
End If

If Val(txtCodLogr.Text) > 0 Then
    Sql = Sql & " AND CODLOGR=" & Val(txtCodLogr.Text)
End If
Sql = Sql & " ORDER BY processogti.ANO,processogti.NUMERO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !CodCidadao > 0 Then
            ReDim Preserve aCodigo(UBound(aCodigo) + 1)
            aCodigo(UBound(aCodigo)) = !CodCidadao
        End If
       .MoveNext
    Loop
   .Close
End With

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

xId = 1
For x = 1 To UBound(aCodigo)
    Sql = "SELECT cidadao.codcidadao, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, cidadao.siglauf, cidade.desccidade, "
    Sql = Sql & "bairro.descbairro, cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,cidadao.nomecidadao,"
    Sql = Sql & "cidadao.cep, cidadao.nomelogradouro AS Rua FROM cidadao LEFT OUTER JOIN cidade ON "
    Sql = Sql & "cidadao.siglauf = cidade.siglauf AND cidadao.codcidade = cidade.codcidade LEFT OUTER JOIN bairro ON cidadao.siglauf = bairro.siglauf AND "
    Sql = Sql & "cidadao.codcidade = bairro.codcidade AND cidadao.codbairro = bairro.codbairro LEFT OUTER JOIN vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO "
    Sql = Sql & "Where Cidadao.CodCidadao = " & aCodigo(x)
    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux3
        sCodInscricao = Format(!CodCidadao, "000000")
        sContribuinte = !nomecidadao
        If IsNull(!NomeLogradouro) Then
            sEnd = !Rua & CStr(SubNull(!NUMIMOVEL))
        Else
            sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(SubNull(!NUMIMOVEL))
        End If
        sCompl = SubNull(!Complemento)
    '    If IsNull(!DescBairro) Then
    '        sBairro = SubNull(!NOMEBairro)
    '    Else
            sBairro = SubNull(!DescBairro)
     '   End If
    '    If IsNull(!desccidade) Then
     '       sCidade = SubNull(!NomeCidade)
      '  Else
            sCidade = SubNull(!desccidade)
       ' End If
        sCEP = SubNull(!Cep)
        sUF = SubNull(!siglauf)
        If sCidade = "JABOTICABAL" And !CodLogradouro > 0 Then
            sCEP = RetornaCEP(!CodLogradouro, !NUMIMOVEL)
        End If
        .Close
    End With
    sCompl = SubNull(Left(sCompl, 20))

    sEndEntrega = sEnd
    sBairroEntrega = sBairro
    sCidEntrega = sCidade
    sCepEntrega = sCEP
    sComplEntrega = sCompl
    sUFEntrega = sUF
    
    
    Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
    Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & Mask(sContribuinte) & "','"
    Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
    cn.Execute Sql, rdExecDirect
    xId = xId + 1
PROXIMO2:
   
Next
frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


End Sub

Private Sub cmdPrint_Click()
Dim sTexto1 As String, sTexto2 As String, sTexto3 As String, sTexto4 As String, sTexto5 As String
Dim x As Integer, sAssunto As String

If IsDate(mskDataDe.Text) And Not IsDate(mskDataAte.Text) Then
   MsgBox "Data inicial sem data final.", vbExclamation, "Atenção"
    Exit Sub
End If

sAssunto = "": txtAssunto.Text = ""
For x = 1 To lvDDList.ListItems.Count
    If lvDDList.ListItems(x).Checked = True Then
        sAssunto = sAssunto & CStr(Val(lvDDList.ListItems(x).SubItems(1))) & ","
    End If
Next

If sAssunto = "" Then
    MsgBox "Selecione um Assunto.", vbExclamation, "Atenção"
    Exit Sub
End If
sAssunto = Left$(sAssunto, Len(sAssunto) - 1)
txtAssunto.Text = sAssunto
If chkTramite.Value = 1 Then
    frmReport.ShowReport "TRAMITEABERTOLOCAL", frmMdi.hwnd, Me.hwnd
    Exit Sub
End If

If Not IsDate(mskDataDe.Text) Then
   MsgBox "Data inicial inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataAte.Text) Then
   MsgBox "Data Final inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If CDate(mskDataDe.Text) > CDate(mskDataAte.Text) Then
   MsgBox "Data inicial maior que a final.", vbExclamation, "Atenção"
   Exit Sub
End If
             
Sql = "DELETE FROM RESUMODIARIO WHERE USUARIO='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
If cmbArquivado.ListIndex = 1 Then
    Sql = "SELECT processogti.ANO, processogti.NUMERO, processogti.DATAENTRADA, processogti.FISICO, processogti.INTERNO, processogti.CODASSUNTO,"
    Sql = Sql & "processogti.DATAARQUIVA, processogti.COMPLEMENTO, cidadao.nomecidadao, processogti.DATAENTRADA AS Expr1, processogti.CENTROCUSTO,"
    Sql = Sql & "processogti.CodCidadao , centrocusto.DESCRICAO, assunto.NOME, vwPROCESSOENDERECO.CodLogr FROM processogti INNER JOIN "
    Sql = Sql & "assunto ON processogti.CODASSUNTO = assunto.CODIGO LEFT OUTER JOIN vwPROCESSOENDERECO ON processogti.ANO = vwPROCESSOENDERECO.ANO AND "
    Sql = Sql & "processogti.NUMERO = vwPROCESSOENDERECO.NUMPROCESSO LEFT OUTER JOIN centrocusto ON processogti.CENTROCUSTO = centrocusto.CODIGO LEFT OUTER JOIN "
    Sql = Sql & "cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
    Sql = Sql & "WHERE DATAENTRADA BETWEEN '" & Format(mskDataDe.Text, "mm/dd/yyyy") & "' AND  '" & Format(mskDataAte.Text, "mm/dd/yyyy") & "' AND CODASSUNTO in (" & sAssunto & ") AND DATAARQUIVA IS NOT NULL"
ElseIf cmbArquivado.ListIndex = 2 Then
    Sql = "SELECT processogti.ANO, processogti.NUMERO, processogti.DATAENTRADA, processogti.FISICO, processogti.INTERNO, processogti.CODASSUNTO,"
    Sql = Sql & "processogti.DATAARQUIVA, processogti.COMPLEMENTO, cidadao.nomecidadao, processogti.DATAENTRADA AS Expr1, processogti.CENTROCUSTO,"
    Sql = Sql & "processogti.CodCidadao , centrocusto.DESCRICAO, assunto.NOME, vwPROCESSOENDERECO.CodLogr FROM processogti INNER JOIN "
    Sql = Sql & "assunto ON processogti.CODASSUNTO = assunto.CODIGO LEFT OUTER JOIN vwPROCESSOENDERECO ON processogti.ANO = vwPROCESSOENDERECO.ANO AND "
    Sql = Sql & "processogti.NUMERO = vwPROCESSOENDERECO.NUMPROCESSO LEFT OUTER JOIN centrocusto ON processogti.CENTROCUSTO = centrocusto.CODIGO LEFT OUTER JOIN "
    Sql = Sql & "cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
    Sql = Sql & "WHERE DATAENTRADA BETWEEN '" & Format(mskDataDe.Text, "mm/dd/yyyy") & "' AND  '" & Format(mskDataAte.Text, "mm/dd/yyyy") & "' AND CODASSUNTO in (" & sAssunto & ") AND DATAARQUIVA IS NULL"
ElseIf cmbArquivado.ListIndex = 0 Then
    Sql = "SELECT processogti.ANO, processogti.NUMERO, processogti.DATAENTRADA, processogti.FISICO, processogti.INTERNO, processogti.CODASSUNTO,"
    Sql = Sql & "processogti.DATAARQUIVA, processogti.COMPLEMENTO, cidadao.nomecidadao, processogti.DATAENTRADA AS Expr1, processogti.CENTROCUSTO,"
    Sql = Sql & "processogti.CodCidadao , centrocusto.DESCRICAO, assunto.NOME, vwPROCESSOENDERECO.CodLogr FROM processogti INNER JOIN "
    Sql = Sql & "assunto ON processogti.CODASSUNTO = assunto.CODIGO LEFT OUTER JOIN vwPROCESSOENDERECO ON processogti.ANO = vwPROCESSOENDERECO.ANO AND "
    Sql = Sql & "processogti.NUMERO = vwPROCESSOENDERECO.NUMPROCESSO LEFT OUTER JOIN centrocusto ON processogti.CENTROCUSTO = centrocusto.CODIGO LEFT OUTER JOIN "
    Sql = Sql & "cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
    Sql = Sql & "WHERE DATAENTRADA BETWEEN '" & Format(mskDataDe.Text, "mm/dd/yyyy") & "' AND  '" & Format(mskDataAte.Text, "mm/dd/yyyy") & "' AND CODASSUNTO in (" & sAssunto & ") "
End If
If Val(txtCodLogr.Text) > 0 Then
    Sql = Sql & " AND CODLOGR=" & Val(txtCodLogr.Text)
End If
Sql = Sql & " ORDER BY ANO,NUMERO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       Do Until .EOF
            If Not IsNull(!nomecidadao) Then
                sTexto1 = !nomecidadao
            Else
                If Not IsNull(!DESCRICAO) Then
                        sTexto1 = !DESCRICAO
                Else
                        sTexto1 = ""
                End If
            End If
            On Error Resume Next
            Sql = "INSERT RESUMODIARIO(USUARIO,NUMPROCESSO,DATAENTRADA,CODASSUNTO,ASSUNTO,REQUERENTE,FISICO,INTERNO,DATADE,DATAATE,COMPLEMENTO,DATAARQUIVA,ANOPROC,NUMPROC) VALUES('"
            Sql = Sql & NomeDoUsuario & "','" & CStr(!Numero) & "-" & RetornaDVProcesso(!Numero) & "/" & CStr(!Ano) & "','" & Format(!DATAENTRADA, "mm/dd/yyyy") & "'," & !CODASSUNTO & ",'" & Mask(Left$(!Complemento, 150)) & "','"
            Sql = Sql & Mask(Left$(Mask(sTexto1), 30)) & "','" & IIf(!FISICO = True, "S", "N") & "','" & IIf(!INTERNO = True, "S", "N") & "','" & CStr(mskDataDe.Text) & "','" & CStr(mskDataAte.Text) & "','" & SubNull(!nome) & "','"
            Sql = Sql & Format(!DATAARQUIVA, "mm/dd/yyyy") & "'," & !Ano & "," & !Numero & ")"
            cn.Execute Sql, rdExecDirect
            DoEvents
            .MoveNext
       Loop
      .Close
End With

If cmbOrder.ListIndex = 0 Then
    frmReport.ShowReport "RESUMOPROTOCOLO", frmMdi.hwnd, Me.hwnd
Else
    frmReport.ShowReport "RESUMOPROTOCOLOREQ", frmMdi.hwnd, Me.hwnd
End If

End Sub


Private Sub cmdReq_Click()
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frmCnsCidadao.show
frmCnsCidadao.ZOrder 0

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
LoadMultiColumnCombo
CheckAll

Sql = "SELECT CODIGO,DESCRICAO FROM CENTROCUSTO ORDER BY DESCRICAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        Do Until .EOF
             cmbSetor.AddItem !DESCRICAO
             cmbSetor.ItemData(cmbSetor.NewIndex) = !Codigo
            .MoveNext
        Loop
       .Close
End With

cmdDDList.Enabled = False
cmbOrder.ListIndex = 0
cmbArquivado.ListIndex = 0
End Sub

Private Sub mskDataAte_GotFocus()
mskDataAte.SetFocus
End Sub

Private Sub mskDataDe_GotFocus()
mskDataDe.SetFocus
End Sub

Private Sub LoadMultiColumnCombo()
Dim itmX As ListItem

Sql = "SELECT CODIGO,NOME,ATIVO FROM ASSUNTO ORDER BY NOME"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvDDList.ListItems.Add(, "C" & Format(!Codigo, "0000"), !nome)
        itmX.SubItems(1) = Format(!Codigo, "0000")
        itmX.SubItems(2) = IIf(!Ativo = True, "Sim", "Não")
       .MoveNext
    Loop
    .Close
End With

End Sub

Private Sub CheckAll()
Dim itemX As ListItem, x As Integer

For x = 1 To lvDDList.ListItems.Count
    lvDDList.ListItems(x).Checked = True
Next

End Sub

Private Sub CheckNone()
Dim itemX As ListItem, x As Integer

For x = 1 To lvDDList.ListItems.Count
    lvDDList.ListItems(x).Checked = False
Next

End Sub


Private Sub txtNomeLogr_Change()
If Trim$(txtNomeLogr) = "" Then
   txtCodLogr.Text = 0
End If
End Sub

Private Sub txtNomeLogr_GotFocus()
txtNomeLogr.SelStart = 0
txtNomeLogr.SelLength = Len(txtNomeLogr)
End Sub

Private Sub txtNomeLogr_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLogr.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLogr) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder (0)
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro não encontrado.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtNomeLogr.SetFocus
          End If
      End With
   End If
Else
   txtCodLogr.Text = 0
End If

End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
   txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   txtCodLogr_LostFocus
   lstNomeLog.Visible = False
   txtNumImovel.SetFocus
End If

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
   txtNomeLogr.SetFocus
End If

End Sub

Private Sub txtCodLogr_LostFocus()
If Val(txtCodLogr.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLogr.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          txtNomeLogr.Text = ""
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtCodLogr.SetFocus
       End If
      .Close
   End With
End If

End Sub


