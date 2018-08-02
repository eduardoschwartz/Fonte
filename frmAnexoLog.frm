VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAnexoLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexo dos processos"
   ClientHeight    =   5610
   ClientLeft      =   8205
   ClientTop       =   5865
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvAnexos 
      Height          =   2115
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   3731
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Nº Processo"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Requerente"
         Object.Width           =   5363
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Assunto"
         Object.Width           =   5363
      EndProperty
   End
   Begin prjChameleon.chameleonButton btAddAnexo 
      Height          =   345
      Left            =   270
      TabIndex        =   3
      ToolTipText     =   "Inserir um novo anexo"
      Top             =   2610
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Incluir um Anexo"
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmAnexoLog.frx":0000
      PICN            =   "frmAnexoLog.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btDelAnexo 
      Height          =   345
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Remover o anexo selecionado"
      Top             =   2610
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Remover Anexo"
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmAnexoLog.frx":0176
      PICN            =   "frmAnexoLog.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvHist 
      Height          =   2115
      Left            =   90
      TabIndex        =   6
      Top             =   3360
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   3731
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Numero"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Nº Processo"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ocorrência"
         Object.Width           =   2365
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Funcionário"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Histórico de ocorrências:"
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   3090
      Width           =   1965
   End
   Begin VB.Label lblNumProc 
      Caption         =   "00000-0/0000"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2730
      TabIndex        =   1
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Processos anexados ao processo:"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2505
   End
End
Attribute VB_Name = "frmAnexoLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btAddAnexo_Click()
Dim z As Variant, nAno As Integer, nNumero As Integer, Sql As String, RdoAux As rdoResultset
Dim nAnoAtual As Integer, nNumeroAtual As Integer, sNumProc As String

z = InputBox("Digite o nº do processo (com DV)", "Incluir Anexo")
If Len(z) > 6 Then
    sNumProc = RetornaNumero(CStr(z))
    nAno = Val(Right(RetornaNumero(sNumProc), 4))
    nNumero = Val(Left$(sNumProc, Len(sNumProc) - 5))
    If nAno = 0 Or nNumero = 0 Then
        MsgBox "Processo inválido", vbCritical, "Erro"
    Else
        nAnoAtual = Val(frmProcesso.lblAno.Caption)
        nNumeroAtual = Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
        If nAno = nAnoAtual And nNumero = nNumeroAtual Then
            MsgBox "Não é possível anexar o mesmo processo.", vbCritical, "Erro"
        Else
            Sql = "select * from processogti where ano=" & nAno & " and numero=" & nNumero
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount = 0 Then
                    MsgBox "processo não cadastrado.", vbCritical, "Erro"
                    RdoAux.Close
                    Exit Sub
                End If
               .Close
            End With
            
            Set itmX = lvHist.ListItems.Add(, , nAno)
            itmX.SubItems(1) = CStr(nNumero)
            itmX.SubItems(2) = Format(Now, "dd/mm/yyyy")
            itmX.SubItems(3) = CStr(nNumero) & "-" & RetornaDVProcesso(CStr(nNumero)) & "/" & CStr(nAno)
            itmX.SubItems(4) = "Incluído"
            itmX.SubItems(5) = RetornaUsuarioFullName
            
            
            Sql = "insert anexo(ano,numero,anoanexo,numeroanexo) values(" & nAnoAtual & "," & nNumeroAtual & ","
            Sql = Sql & nAno & "," & nNumero & ")"
            cn.Execute Sql, rdExecDirect
            
            Sql = "insert anexo_log (ano,numero,ano_anexo,numero_anexo,removido,data,userid) values("
            Sql = Sql & nAnoAtual & "," & nNumeroAtual & "," & nAno & "," & nNumero & "," & 0 & ",'"
            Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
            cn.Execute Sql, rdExecDirect
            
        End If
    End If
Else
    MsgBox "Processo inválido", vbCritical, "Erro"
End If

End Sub

Private Sub btDelAnexo_Click()
Dim z As Variant, nAno As Integer, nNumero As Integer, Sql As String, nAnoAnexo As Integer, nNumeroAnexo As Integer

If lvAnexos.ListItems.Count = 0 Then Exit Sub
If lvAnexos.SelectedItem = Null Then
    MsgBox "Selecione um anexo.", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Remover este anexo?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    nAno = Val(frmProcesso.lblAno.Caption)
    nNumero = Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
    nAnoAnexo = Val(lvAnexos.SelectedItem.Text)
    nNumeroAnexo = Val(lvAnexos.SelectedItem.SubItems(1))
    
    Sql = "DELETE FROM ANEXO WHERE ANO=" & nAno & " AND NUMERO=" & nNumero & " AND ANOANEXO=" & nAnoAnexo & " AND NUMEROANEXO=" & nNumeroAnexo
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM ANEXO WHERE ANO=" & nAnoAnexo & " AND NUMERO=" & nNumeroAnexo & " AND ANOANEXO=" & nAno & " AND NUMEROANEXO=" & nNumero
    cn.Execute Sql, rdExecDirect
    
    lvAnexos.ListItems.Remove (lvAnexos.SelectedItem.Index)
    
    Sql = "insert anexo_log (ano,numero,ano_anexo,numero_anexo,removido,data,userid) values("
    Sql = Sql & nAno & "," & nNumero & "," & nAnoAnexo & "," & nNumeroAnexo & "," & 1 & ",'"
    Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
    cn.Execute Sql, rdExecDirect

    Set itmX = lvHist.ListItems.Add(, , nAnoAnexo)
    itmX.SubItems(1) = CStr(nNumeroAnexo)
    itmX.SubItems(2) = Format(Now, "dd/mm/yyyy")
    itmX.SubItems(3) = CStr(nNumeroAnexo) & "-" & RetornaDVProcesso(CStr(nNumeroAnexo)) & "/" & CStr(nAnoAnexo)
    itmX.SubItems(4) = "Removido"
    itmX.SubItems(5) = RetornaUsuarioFullName
    
End If

End Sub

Private Sub Form_Load()
Dim nAno As Integer, nNumero As Integer

Centraliza Me

lblNumProc.Caption = frmProcesso.lblNumProc.Caption & "/" & frmProcesso.lblAno.Caption
nAno = Val(frmProcesso.lblAno.Caption)
nNumero = Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))

CarregaAnexo nNumero, nAno

End Sub

Private Sub CarregaAnexo(Numero As Integer, Ano As Integer)

Dim Sql As String, RdoAux As rdoResultset

Sql = "SELECT anexo.ano, anexo.numero, anexo.anoanexo, anexo.numeroanexo, vwFULLPROCESSO.nomecidadao,vwFULLPROCESSO.descricao, "
Sql = Sql & "vwFULLPROCESSO.Complemento FROM anexo INNER JOIN vwFULLPROCESSO ON anexo.anoanexo = vwFULLPROCESSO.ANO AND anexo.numeroanexo = vwFULLPROCESSO.NUMERO where anexo.ano=" & Ano & " and anexo.numero=" & Numero
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvAnexos.ListItems.Add(, , !anoanexo)
        itmX.SubItems(1) = CStr(!numeroanexo)
        itmX.SubItems(2) = CStr(!numeroanexo) & "-" & RetornaDVProcesso(!numeroanexo) & "/" & CStr(!anoanexo)
        itmX.SubItems(3) = IIf(IsNull(!nomecidadao), SubNull(!descricao), !nomecidadao)
        itmX.SubItems(4) = !Complemento
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT anexo_log.sid, anexo_log.ano, anexo_log.numero, anexo_log.ano_anexo, anexo_log.numero_anexo, anexo_log.removido, anexo_log.data, anexo_log.userid, vwFULLPROCESSO.nomecidadao, "
Sql = Sql & "vwFULLPROCESSO.Complemento , USUARIO.NomeCompleto FROM anexo_log INNER JOIN vwFULLPROCESSO ON anexo_log.ano_anexo = vwFULLPROCESSO.ANO AND anexo_log.numero_anexo = vwFULLPROCESSO.NUMERO INNER JOIN "
Sql = Sql & "usuario ON anexo_log.userid = usuario.Id where anexo_log.ano=" & Ano & " and anexo_log.numero=" & Numero
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvHist.ListItems.Add(, , !ano_anexo)
        itmX.SubItems(1) = CStr(!numero_anexo)
        itmX.SubItems(2) = Format(!Data, "dd/mm/yyyy")
        itmX.SubItems(3) = CStr(!numero_anexo) & "-" & RetornaDVProcesso(!numero_anexo) & "/" & CStr(!ano_anexo)
        itmX.SubItems(4) = IIf(!removido, "Removido", "Incluído")
        itmX.SubItems(5) = Mask(!NomeCompleto)
       .MoveNext
    Loop
   .Close
End With

End Sub
