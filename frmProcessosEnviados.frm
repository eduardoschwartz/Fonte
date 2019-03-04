VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmProcessosEnviados 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de processos enviados por centro de custos"
   ClientHeight    =   1620
   ClientLeft      =   14475
   ClientTop       =   8985
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   6450
   Begin VB.CheckBox chkSemDataEnvio 
      Caption         =   "Sem data de envio"
      Height          =   195
      Left            =   3825
      TabIndex        =   8
      Top             =   720
      Width           =   2220
   End
   Begin VB.ComboBox cmbSetor 
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4605
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3810
      TabIndex        =   3
      ToolTipText     =   "Imprimir Relatório"
      Top             =   1080
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
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
      MICON           =   "frmProcessosEnviados.frx":0000
      PICN            =   "frmProcessosEnviados.frx":001C
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
      Left            =   4950
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Sair"
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
      MICON           =   "frmProcessosEnviados.frx":0176
      PICN            =   "frmProcessosEnviados.frx":0192
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
      Left            =   1485
      TabIndex        =   1
      Top             =   765
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmProcessosEnviados.frx":0200
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
   Begin esMaskEdit.esMaskedEdit mskData2 
      Height          =   285
      Left            =   1485
      TabIndex        =   2
      Top             =   1125
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmProcessosEnviados.frx":021C
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
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Data final:"
      Height          =   195
      Index           =   0
      Left            =   495
      TabIndex        =   7
      Top             =   1185
      Width           =   795
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Data inicio:"
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   6
      Top             =   825
      Width           =   795
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Setor........:"
      Height          =   195
      Index           =   3
      Left            =   450
      TabIndex        =   5
      Top             =   315
      Width           =   840
   End
End
Attribute VB_Name = "frmProcessosEnviados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset

Private Sub cmdPrint_Click()
Dim nCCusto As Integer, nSeq As Integer, nAno As Integer, nNumero As Long, RdoAux2 As rdoResultset, sDesc1 As String, sDesc2 As String
Dim sNome1 As String, sNome2 As String, RdoAux3 As rdoResultset, sRequerente As String, sAssunto As String

'If chkSemDataEnvio.Value = vbUnchecked Then
    If Not IsDate(mskData.Text) Then
        MsgBox "Data inicial inválida", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    If Not IsDate(mskData2.Text) Then
        MsgBox "Data final inválida", vbExclamation, "Atenção"
        Exit Sub
    End If
'End If
Sql = "DELETE FROM PROCESSOENVIO WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

nCCusto = cmbSetor.ItemData(cmbSetor.ListIndex)
Ocupado
If chkSemDataEnvio.value = vbChecked Then
    Sql = "SELECT * From vwTRAMITACAO2 WHERE datahora between '" & Format(mskData.Text, "mm/dd/yyyy") & "' and '" & Format(mskData2.Text, "mm/dd/yyyy") & "' AND ccusto = " & nCCusto & " ORDER BY ANO,NUMERO"
Else
    Sql = "SELECT * From vwTRAMITACAO2 WHERE dataenvio between '" & Format(mskData.Text, "mm/dd/yyyy") & "' and '" & Format(mskData2.Text, "mm/dd/yyyy") & "' AND ccusto = " & nCCusto
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        nAno = !Ano
        nNumero = !Numero
        
  '      If nNumero = 14726 Then MsgBox "teste"
        
        Sql = "SELECT * FROM PROCESSOGTI WHERE ANO=" & nAno & " AND NUMERO=" & nNumero
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        sAssunto = Mask(Left(RdoAux2!Complemento, 50))
        If RdoAux2!ORIGEM = 1 Then 'CENTRO CUSTO
            If RdoAux2.RowCount > 0 Then
                Sql = "SELECT DESCRICAO FROM CENTROCUSTO WHERE CODIGO=" & RdoAux2!CENTROCUSTO
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux3.RowCount > 0 Then
                    sRequerente = RdoAux3!descricao
                Else
                    GoTo cidadao
                End If
                RdoAux3.Close
            Else
                GoTo cidadao
            End If
        Else 'CIDADAO
cidadao:
            Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & RdoAux2!CodCidadao
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
            If RdoAux3.RowCount > 0 Then
                sRequerente = RdoAux3!NomeCidadao
            Else
                sRequerente = ""
            End If
            RdoAux3.Close
        End If
        
                
        nSeq = !Seq
        sDesc1 = !descricao
        sNome1 = SubNull(!NomeCompleto)
        Sql = "SELECT * FROM vwTRAMITACAO2 WHERE ANO=" & !Ano & " AND NUMERO=" & !Numero & " AND SEQ=" & nSeq + 1
        'Sql = "SELECT * FROM vwTRAMITACAO2 WHERE ANO=" & !Ano & " AND NUMERO=" & !Numero & " AND SEQ=" & nSeq
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        
        With RdoAux2
            If .RowCount > 0 Then
                sDesc2 = !descricao
                sNome2 = SubNull(!NomeCompleto)
            Else
                Sql = "SELECT tramitacaocc.ano, tramitacaocc.numero, tramitacaocc.seq, tramitacaocc.ccusto, centrocusto.DESCRICAO "
                Sql = Sql & "FROM tramitacaocc INNER JOIN centrocusto ON tramitacaocc.ccusto = centrocusto.CODIGO "
                Sql = Sql & "Where tramitacaocc.Ano = " & nAno & " And tramitacaocc.Numero = " & nNumero & " AND SEQ=" & nSeq + 1
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux2.RowCount > 0 Then
                    sDesc2 = RdoAux2!descricao
                    sNome2 = ""
                Else
                    GoTo proximo
                End If
            End If
           .Close
        End With
        On Error Resume Next
        Sql = "INSERT PROCESSOENVIO(COMPUTER,ANO,NUMERO,PROCESSO,SEQ,DESC1,NOME1,DESC2,NOME2,DATAENVIO,ASSUNTO,REQUERENTE,DATAENTRADA) VALUES('" & NomeDeLogin & "',"
        Sql = Sql & nAno & "," & nNumero & ",'" & CStr(nNumero) & "-" & CStr(RetornaDVProcesso(CLng(nNumero))) & "/" & CStr(nAno) & "'," & nSeq & ",'" & Mask(sDesc1) & "','" & Mask(sNome1) & "','" & Mask(sDesc2) & "','" & Mask(sNome2) & "','"
        Sql = Sql & Format(!DATAENVIO, "mm/dd/yyyy") & "','" & Left(Mask(sAssunto), 50) & "','" & Left(Mask(sRequerente), 50) & "','" & Format(!DATAHORA, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
        On Error GoTo 0
proximo:
       .MoveNext
    Loop
   .Close
End With
Liberado
If chkSemDataEnvio.value = vbChecked Then
    frmReport.ShowReport2 "PROCESSOENVIADO2", frmMdi.HWND, Me.HWND
Else
    frmReport.ShowReport "PROCESSOENVIADO", frmMdi.HWND, Me.HWND
End If

Sql = "DELETE FROM PROCESSOENVIO WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Sql = "SELECT CODIGO,DESCRICAO FROM CENTROCUSTO where ativo=1 ORDER BY DESCRICAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        Do Until .EOF
             cmbSetor.AddItem !descricao
             cmbSetor.ItemData(cmbSetor.NewIndex) = !Codigo
            .MoveNext
        Loop
       .Close
End With

Centraliza Me
End Sub

Private Sub mskData_GotFocus()
mskData.SetFocus
End Sub

Private Sub mskData2_GotFocus()
mskData2.SetFocus
End Sub
