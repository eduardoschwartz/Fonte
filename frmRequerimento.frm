VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmRequerimento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimento para abertura de processo"
   ClientHeight    =   5880
   ClientLeft      =   2655
   ClientTop       =   2370
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9975
   Begin VB.TextBox txtEndereco 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1395
      TabIndex        =   18
      Top             =   810
      Width           =   8385
   End
   Begin VB.OptionButton optEnd 
      Caption         =   "Endereço Com."
      Height          =   195
      Index           =   1
      Left            =   8415
      TabIndex        =   17
      Top             =   495
      Width           =   1455
   End
   Begin VB.OptionButton optEnd 
      Caption         =   "Endereço Res."
      Height          =   195
      Index           =   0
      Left            =   6930
      TabIndex        =   16
      Top             =   495
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton OptP 
      Caption         =   "Pessoa Jurídica"
      Height          =   240
      Index           =   1
      Left            =   1665
      TabIndex        =   15
      Top             =   5535
      Width           =   1545
   End
   Begin VB.OptionButton OptP 
      Caption         =   "Pessoa Física"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   14
      Top             =   5535
      Width           =   1365
   End
   Begin VB.TextBox txtEnd 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   90
      MaxLength       =   5000
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4500
      Width           =   9780
   End
   Begin VB.TextBox txtReq 
      Appearance      =   0  'Flat
      Height          =   2985
      Left            =   90
      MaxLength       =   5000
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1170
      Width           =   9780
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   9270
      TabIndex        =   0
      ToolTipText     =   "Consulta Cidadão"
      Top             =   90
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
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
      MICON           =   "frmRequerimento.frx":0000
      PICN            =   "frmRequerimento.frx":001C
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
      Height          =   315
      Left            =   7650
      TabIndex        =   11
      ToolTipText     =   "Imprimir Requerimento"
      Top             =   5490
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmRequerimento.frx":0176
      PICN            =   "frmRequerimento.frx":0192
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
      Left            =   8730
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   5490
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
      MICON           =   "frmRequerimento.frx":02EC
      PICN            =   "frmRequerimento.frx":0308
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
      Caption         =   "Endereço do Imóvel/Estabelecimento"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   135
      TabIndex        =   13
      Top             =   4275
      Width           =   5610
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   135
      TabIndex        =   8
      Top             =   810
      Width           =   1200
   End
   Begin VB.Label lblCPF 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4500
      TabIndex        =   7
      Top             =   495
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "CPF/CNPJ.:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3375
      TabIndex        =   6
      Top             =   495
      Width           =   1110
   End
   Begin VB.Label lblRG 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1395
      TabIndex        =   5
      Top             =   495
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Nº DE RG..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   495
      Width           =   1200
   End
   Begin VB.Label lblCodRequerente 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   3
      Top             =   180
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Requerente (Cidadão)..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   180
      Width           =   2505
   End
   Begin VB.Label lblRequerente 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3465
      TabIndex        =   1
      Top             =   180
      Width           =   5745
   End
End
Attribute VB_Name = "frmRequerimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCnsImovel_Click()
Set frm = frmCidadao
frm.sForm = Me.Name
frm.show
frm.ZOrder 0
End Sub

Private Sub cmdPrint_Click()
Dim Sql As String

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

If Val(lblCodRequerente.Caption) > 0 Then
    Sql = "INSERT REPORTTMP(USUARIO,MEMO1,MEMO2) VALUES('" & NomeDeLogin & "','" & Mask(txtReq.Text) & "','" & Mask(txtEnd.Text) & "')"
    cn.Execute Sql, rdExecDirect
    frmReport.ShowReport "REQUERIMENTOPROC", frmMdi.hwnd, Me.hwnd

    Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Val(lblCodRequerente.Caption) = 0 Then Exit Sub
If Val(lblCodRequerente.Caption) > 500000 Then
    Le
Else
    MsgBox "Código de cidadão inválido.", vbExclamation, "Atenção"
    Limpa
End If

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset

If optEnd(0).Value = True Then
    Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
    Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
    Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
    Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(lblCodRequerente.Caption)
Else
    Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
    Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
    Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
    Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(lblCodRequerente.Caption)
End If
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
On Error Resume Next
With RdoAux2
    If .RowCount > 0 Then
        lblRG.Caption = SubNull(!frg) & " " & SubNull(!fORGAO)
        lblCPF.Caption = SubNull(!CPF)
        If lblCPF.Caption <> "" Then
            lblCPF.Caption = Format(!CPF, "00#\.###\.###-##")
        End If
        If lblCPF.Caption = "" Then
            If Not IsNull(!Cnpj) Then
                lblCPF.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
            End If
        End If
                 
        If Val(SubNull(!FCodLogradouro)) > 0 Then
            Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
            Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
            Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
            Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoS
                If .RowCount > 0 Then
                   sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                Else
                   sEnd = ""
                End If
                .Close
            End With
        Else
           sEnd = SubNull(!FNomeLogradouro)
        End If
        sEnd = sEnd & " " & SubNull(RdoAux2!fNUMIMOVEL) & " " & SubNull(!fcomplemento)
          
        Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
        Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
        If RdoS.RowCount > 0 Then
            sCidade = "na cidade de " & RdoS!desccidade
        Else
             sCidade = ""
        End If
        
        If Not IsNull(!CodBairro) Then
            Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
            If .RowCount > 0 Then
                sBairro = RdoS!DescBairro
            Else
                sBairro = ""
            End If
        Else
            sBairro = ""
        End If
        sUF = SubNull(!fsiglauf)
        sFone = SubNull(!TELEFONE)
        sCEP = SubNull(!FCEP)
        sEnd = sEnd & " " & SubNull(sBairro) & " " & SubNull(sCidade)
        
    Else
        sEnd = ""
        sBairro = ""
        sCidade = ""
        sFone = ""
        sUF = ""
        sCEP = ""
    End If
   .Close
End With

txtEndereco.Text = sEnd

End Sub

Private Sub Limpa()
lblCodRequerente.Caption = ""
lblRequerente.Caption = ""
lblRG.Caption = ""
lblCPF.Caption = ""
lblEndereco.Caption = ""
End Sub

Private Sub lblEndereco_Click()

End Sub
