VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmMalaDiretaISS 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mala Direta para ISS"
   ClientHeight    =   2865
   ClientLeft      =   5520
   ClientTop       =   4095
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   5910
   Begin Tributacao.XP_ProgressBar Pbar 
      Height          =   240
      Left            =   2835
      TabIndex        =   9
      Top             =   1485
      Width           =   2715
      _ExtentX        =   4789
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
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "Carregar Senhas"
      Top             =   2400
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Senhas"
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
      MICON           =   "frmMalaDiretaISS.frx":0000
      PICN            =   "frmMalaDiretaISS.frx":001C
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
      Height          =   315
      Left            =   4650
      TabIndex        =   2
      ToolTipText     =   "Sair da Tela"
      Top             =   675
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmMalaDiretaISS.frx":0176
      PICN            =   "frmMalaDiretaISS.frx":0192
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4650
      TabIndex        =   3
      ToolTipText     =   "Cancelar Edição"
      Top             =   270
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gerar"
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
      MICON           =   "frmMalaDiretaISS.frx":0200
      PICN            =   "frmMalaDiretaISS.frx":021C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Tipo de Tributo"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2145
      Begin VB.OptionButton opt 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Atividade ISS"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1500
         Width           =   1755
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00EEEEEE&
         Caption         =   "IPTU"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   7
         Top             =   1125
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Diversos"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   750
         Width           =   1755
      End
      Begin VB.OptionButton opt 
         BackColor       =   &H00EEEEEE&
         Caption         =   "ISS Variável"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   390
         Width           =   1755
      End
   End
   Begin prjChameleon.chameleonButton cmdEtiqueta 
      Height          =   345
      Left            =   3300
      TabIndex        =   4
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   2400
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
      MICON           =   "frmMalaDiretaISS.frx":02BB
      PICN            =   "frmMalaDiretaISS.frx":02D7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmMalaDiretaISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEtiqueta_Click()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim xId As Long, nNumRec As Long, nCodLogr As Long, sCodInscricao As String, sContribuinte As String
Dim sEnd As String, nNum As Integer, sCEP As String, sCompl As String, sBairro As String
Dim sEndEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String
Dim z As Variant

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


If Opt(0).Value = True Then
    xId = 1
    Sql = "SELECT DISTINCT CODREDUZIDO From SENHACONSIST where CODREDUZIDO<500000"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        CallPb xId, nNumRec
        Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
        Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,"
        Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
        Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
        Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
        Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                GoTo Proximo
            End If
            nCodLogr = !CodLogradouro
            sCodInscricao = Format(!codigomob, "000000")
            sContribuinte = !razaosocial
            Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & nCodLogr
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero)
                    nNum = RdoAux2!Numero
                Else
                    nNum = 0
                End If
               .Close
            End With
            sCEP = RetornaCEP(nCodLogr, nNum)
            sCompl = SubNull(Left(!Complemento, 20))
            sBairro = SubNull(!DescBairro)
    
            sEndEntrega = sEnd
            sBairroEntrega = sBairro
            sCidEntrega = "JABOTICABAL"
            sCepEntrega = sCEP
            sComplEntrega = sCompl
            sUFEntrega = "SP"
            
           .Close
        End With
        
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
    'Exit Sub
    
    xId = 1
    Sql = "SELECT DISTINCT CODREDUZIDO From SENHACONSIST where CODREDUZIDO>=500000"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        CallPb xId, nNumRec
        Sql = "SELECT cidadao.codcidadao, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, cidadao.siglauf, cidade.desccidade, "
        Sql = Sql & "bairro.descbairro, cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,cidadao.nomecidadao,"
        Sql = Sql & "cidadao.cep, cidadao.nomelogradouro AS Rua FROM cidadao LEFT OUTER JOIN cidade ON "
        Sql = Sql & "cidadao.siglauf = cidade.siglauf AND cidadao.codcidade = cidade.codcidade LEFT OUTER JOIN bairro ON cidadao.siglauf = bairro.siglauf AND "
        Sql = Sql & "cidadao.codcidade = bairro.codcidade AND cidadao.codbairro = bairro.codbairro LEFT OUTER JOIN vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO "
        Sql = Sql & "Where Cidadao.CodCidadao = " & RdoAux!CODREDUZIDO
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
         '       sBairro = SubNull(!NOMEBairro)
          '  Else
                sBairro = SubNull(!DescBairro)
           ' End If
            'If IsNull(!desccidade) Then
             '   sCidade = SubNull(!NomeCidade)
           ' Else
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
        'sBairro = SubNull(!DescBairro)
    
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
       .MoveNext
        Loop
       .Close
    End With
ElseIf Opt(3).Value = True Then 'ISS VARIAVEL
    z = InputBox("Digite o código da atividade.")
    If Val(z) = 0 Then Exit Sub
    Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
    Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,"
    Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
    Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
    Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
    Sql = Sql & "Where MOBILIARIO.CODATIVIDADE = " & z & " AND DATAENCERRAMENTO IS NULL"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        CallPb xId, nNumRec
        nCodLogr = !CodLogradouro
        sCodInscricao = Format(!codigomob, "000000")
        sContribuinte = !razaosocial
        Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & nCodLogr
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If .RowCount > 0 Then
                sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux!Numero)
                nNum = RdoAux!Numero
            Else
                nNum = 0
            End If
           .Close
        End With
        sCEP = RetornaCEP(nCodLogr, nNum)
        sCompl = SubNull(Left(!Complemento, 20))
        sBairro = SubNull(!DescBairro)
        sEndEntrega = sEnd
        sBairroEntrega = sBairro
        sCidEntrega = "JABOTICABAL"
        sCepEntrega = sCEP
        sComplEntrega = sCompl
        sUFEntrega = "SP"
        
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & Mask(sContribuinte) & "','"
        Sql = Sql & Left(sEndEntrega & " " & sComplEntrega, 60) & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
        cn.Execute Sql, rdExecDirect
        xId = xId + 1
       .MoveNext
        Loop
       .Close
    End With
End If


frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdGerar_Click()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim xId As Long, nNumRec As Long, nCodLogr As Long, sCodInscricao As String, sContribuinte As String
Dim sEnd As String, nNum As Integer, sCEP As String, sCompl As String, sBairro As String
Dim sEndEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

xId = 1
If Opt(0).Value = True Then 'ISS VARIAVEL
    Sql = "SELECT DISTINCT CODREDUZIDO From DEBITOPARCELA "
    Sql = Sql & "WHERE (ANOEXERCICIO = " & Year(Now) & ") AND CODLANCAMENTO = 5 ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        CallPb xId, nNumRec
        Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
        Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,"
        Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
        Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
        Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
        Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                GoTo Proximo
            End If
            nCodLogr = !CodLogradouro
            sCodInscricao = Format(!codigomob, "000000")
            sContribuinte = !razaosocial
            Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & nCodLogr
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero)
                    nNum = RdoAux2!Numero
                Else
                    nNum = 0
                End If
               .Close
            End With
            sCEP = RetornaCEP(nCodLogr, nNum)
            sCompl = SubNull(Left(!Complemento, 20))
            sBairro = SubNull(!DescBairro)

            sEndEntrega = sEnd
            sBairroEntrega = sBairro
            sCidEntrega = "JABOTICABAL"
            sCepEntrega = sCEP
            sComplEntrega = sCompl
            sUFEntrega = "SP"

           
           .Close
        End With
        
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & sContribuinte & "','"
        Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
        cn.Execute Sql, rdExecDirect
        xId = xId + 1
Proximo:
       .MoveNext
        Loop
       .Close
    End With
ElseIf Opt(1).Value = True Then 'DIVERSOS
    Sql = "SELECT CODIGO From TABELATMP"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        CallPb xId, nNumRec
        Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
        Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,NOMELOGRADOURO,"
        Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
        Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
        Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
        Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & !Codigo
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount = 0 Then
                GoTo PROXIMO2
            End If
            nCodLogr = !CodLogradouro
            sCodInscricao = Format(!codigomob, "000000")
            sContribuinte = !razaosocial
            Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & nCodLogr
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If .RowCount > 0 Then
                    sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero)
                    nNum = RdoAux2!Numero
                Else
                    sEnd = SubNull(RdoAux2!NomeLogradouro) & " Nº " & CStr(SubNull(RdoAux2!Numero))
                    nNum = Val(SubNull(RdoAux2!Numero))
                End If
               .Close
            End With
            sCEP = RetornaCEP(nCodLogr, nNum)
            sCompl = SubNull(Left(!Complemento, 20))
            sBairro = SubNull(!DescBairro)

            sEndEntrega = sEnd
            sBairroEntrega = sBairro
            sCidEntrega = !desccidade
            sCepEntrega = sCEP
            sComplEntrega = sCompl
            sUFEntrega = "SP"
            
           .Close
        End With
        
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & Mask(sContribuinte) & "','"
        Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
        cn.Execute Sql, rdExecDirect
        xId = xId + 1
PROXIMO2:
       .MoveNext
        Loop
       .Close
    End With
ElseIf Opt(2).Value = True Then 'IPTU
    Sql = "SELECT * FROM VWFULLIMOVEL WHERE LI_CODBAIRRO=61 ORDER BY NOMECIDADAO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        CallPb xId, nNumRec
            sContribuinte = !nomecidadao
            sEnd = !Logradouro & " Nº " & CStr(!Li_Num)
            sCEP = RetornaCEP(!CodLogr, !Li_Num)
            sCompl = SubNull(Left(!Li_Compl, 20))
            sBairro = SubNull(!DescBairro)

            sEndEntrega = sEnd
            sBairroEntrega = sBairro
            sCidEntrega = "JABOTICABAL"
            sCepEntrega = sCEP
            sComplEntrega = sCompl
            sUFEntrega = "SP"
            
        
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & Mask(sContribuinte) & "','"
        Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
        cn.Execute Sql, rdExecDirect
        xId = xId + 1
proximo3:
       .MoveNext
        Loop
       .Close
    End With
ElseIf Opt(3).Value = True Then 'ISS VARIAVEL
    Sql = "SELECT MOBILIARIO.CODIGOMOB,MOBILIARIO.DVMOB,MOBILIARIO.RAZAOSOCIAL,MOBILIARIO.NOMEFANTASIA,"
    Sql = Sql & "MOBILIARIO.NUMERO,MOBILIARIO.CODLOGRADOURO,"
    Sql = Sql & "MOBILIARIO.COMPLEMENTO,BAIRRO.DESCBAIRRO,CIDADE.DESCCIDADE,MOBILIARIO.CODATIVIDADE,MOBILIARIO.ATIVEXTENSO "
    Sql = Sql & "FROM MOBILIARIO LEFT OUTER JOIN CIDADE ON MOBILIARIO.SIGLAUF = CIDADE.SIGLAUF AND MOBILIARIO.CODCIDADE = CIDADE.CODCIDADE LEFT OUTER JOIN "
    Sql = Sql & "BAIRRO ON MOBILIARIO.SIGLAUF = BAIRRO.SIGLAUF AND MOBILIARIO.CODCIDADE = BAIRRO.CODCIDADE AND MOBILIARIO.CODBAIRRO = BAIRRO.CODBAIRRO "
    Sql = Sql & "Where MOBILIARIO.CODATIVIDADE = 21001"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        CallPb xId, nNumRec
        nCodLogr = !CodLogradouro
        sCodInscricao = Format(!codigomob, "000000")
        sContribuinte = !razaosocial
        Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & nCodLogr
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If .RowCount > 0 Then
                sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & CStr(RdoAux2!Numero)
                nNum = RdoAux2!Numero
            Else
                nNum = 0
            End If
           .Close
        End With
        sCEP = RetornaCEP(nCodLogr, nNum)
        sCompl = SubNull(Left(!Complemento, 20))
        sBairro = SubNull(!DescBairro)
        sEndEntrega = sEnd
        sBairroEntrega = sBairro
        sCidEntrega = "JABOTICABAL"
        sCepEntrega = sCEP
        sComplEntrega = sCompl
        sUFEntrega = "SP"
        
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & sContribuinte & "','"
        Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
        cn.Execute Sql, rdExecDirect
        xId = xId + 1
       .MoveNext
        Loop
       .Close
    End With
End If

frmReport.ShowReport "ETIQUETAPROTOCOLO", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdPrint_Click()
Dim RdoAux As rdoResultset, Sql As String

Sql = "TRUNCATE TABLE SENHACONSIST"
cn.Execute Sql, rdExecDirect

Sql = "SELECT DISTINCT mobiliario.codigomob, mobiliario.razaosocial, senhaconsisttmp.login, senhaconsisttmp.senha FROM senhaconsisttmp INNER JOIN "
Sql = Sql & "mobiliario ON senhaconsisttmp.codreduzido = mobiliario.codigomob Where (mobiliario.codigomob < 500000) ORDER BY mobiliario.codigomob"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
            Sql = "INSERT SENHACONSIST (CODREDUZIDO,RAZAOSOCIAL,LOGIN,SENHA) VALUES(" & !codigomob & ",'" & Left(Mask(!razaosocial), 50) & "','" & !Login & "','" & !SENHA & "')"
            cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, senhaconsisttmp.login, senhaconsisttmp.senha FROM senhaconsisttmp INNER JOIN "
Sql = Sql & "cidadao ON senhaconsisttmp.codreduzido = cidadao.codcidadao  where cidadao.codcidadao>=500000 ORDER BY cidadao.codcidadao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
            Sql = "INSERT SENHACONSIST (CODREDUZIDO,RAZAOSOCIAL,LOGIN,SENHA) VALUES(" & !CodCidadao & ",'" & Left(Mask(!nomecidadao), 50) & "','" & !Login & "','" & !SENHA & "')"
            cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

MsgBox "Senhas importadas.", vbExclamation, "Confirmação"

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If ((nVal * 100) / nTot) <= 100 Then
   PBar.Value = (nVal * 100) / nTot
Else
   PBar.Value = 100
End If
Me.Refresh
If cGetInputState() <> 0 Then DoEvents
End Sub

