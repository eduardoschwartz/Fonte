VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAnexaProc 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexos do Processo"
   ClientHeight    =   4890
   ClientLeft      =   10275
   ClientTop       =   2700
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   9360
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   1545
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2790
      Width           =   9240
   End
   Begin MSFlexGridLib.MSFlexGrid grdAnexo 
      Height          =   1755
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   3096
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmAnexaProc.frx":0000
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   7665
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   4470
      Width           =   1575
      _ExtentX        =   2778
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmAnexaProc.frx":00A8
      PICN            =   "frmAnexaProc.frx":00C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdInserir 
      Height          =   345
      Left            =   45
      TabIndex        =   10
      ToolTipText     =   "Inserir um novo local para a tramitação"
      Top             =   4455
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Inserir Anexo"
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
      MICON           =   "frmAnexaProc.frx":021E
      PICN            =   "frmAnexaProc.frx":023A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRemover 
      Height          =   345
      Left            =   1725
      TabIndex        =   11
      ToolTipText     =   "Remover um local de tramitação"
      Top             =   4470
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frmAnexaProc.frx":0394
      PICN            =   "frmAnexaProc.frx":03B0
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
      Caption         =   "Nº do Processo...:"
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Height          =   225
      Index           =   1
      Left            =   2490
      TabIndex        =   8
      Top             =   90
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto...............:"
      Height          =   225
      Index           =   6
      Left            =   90
      TabIndex        =   7
      Top             =   390
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requerente.........:"
      Height          =   225
      Index           =   7
      Left            =   90
      TabIndex        =   6
      Top             =   690
      Width           =   1365
   End
   Begin VB.Label lblNumProc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1500
      TabIndex        =   5
      Top             =   90
      Width           =   915
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2970
      TabIndex        =   4
      Top             =   90
      Width           =   705
   End
   Begin VB.Label lblAssunto 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1500
      TabIndex        =   3
      Top             =   390
      Width           =   6495
   End
   Begin VB.Label lblRequerente 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1500
      TabIndex        =   2
      Top             =   690
      Width           =   6495
   End
End
Attribute VB_Name = "frmAnexaProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Processo
    nAno As Integer
    nNumproc As Long
End Type

Dim RdoAux As rdoResultset, Sql As String
Dim sDataEntrada As String, sDataCancel As String
Dim sDataArquiva As String, sDataSuspenso As String
Dim bEsp As Boolean, evEsp As String, sRet As String

Private Sub cmdInserir_Click()
frmCnsProcesso2.show
frmCnsProcesso2.ZOrder 0
End Sub

Private Sub cmdRemover_Click()
Dim nNumproc As Long, nAno As Integer, nNumProc2 As Long, nAno2 As Integer, sObs As String, RdoAux2 As rdoResultset, sObsOld As String
If grdAnexo.Rows = 1 Then Exit Sub
nAno = Val(lblAno.Caption)
nNumproc = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
nAno2 = grdAnexo.TextMatrix(grdAnexo.Row, 0)
nNumProc2 = grdAnexo.TextMatrix(grdAnexo.Row, 1)

Sql = "DELETE FROM ANEXO WHERE ANO=" & nAno & " AND NUMERO=" & nNumproc & " AND ANOANEXO=" & nAno2 & " AND NUMEROANEXO=" & nNumProc2
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM ANEXO WHERE ANO=" & nAno2 & " AND NUMERO=" & nNumProc2 & " AND ANOANEXO=" & nAno & " AND NUMEROANEXO=" & nNumproc
cn.Execute Sql, rdExecDirect

With grdAnexo
    If .Rows > 1 Then
        If .Rows > 2 Then
            .RemoveItem (.Row)
        Else
            .Rows = 1
        End If
    End If
End With

sObs = "O processo anexado nº " & nNumProc2 & "-" & RetornaDVProcesso(nNumProc2) & "/" & nAno2 & " foi removido do processo em " & Format(Now, "dd/mm/yyyy") & " por " & RetornaUsuarioFullName2(NomeDeLogin) & "."
'sObs = "O processo anexado nº " & nNumProc2 & "-" & RetornaDVProcesso(nNumProc2) & "/" & nAno2 & " foi removido do processo em " & Format(Now, "dd/mm/yyyy") & " por PROTOCOLO/ARQUIVO."
If txtObs.Text <> "" Then txtObs.Text = txtObs.Text & vbCrLf
txtObs.Text = txtObs.Text & sObs

Sql = "SELECT * FROM PROCESSOGTI where ano=" & nAno & " and numero=" & nNumproc
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
sObsOld = SubNull(RdoAux2!OBSANEXO)
RdoAux2.Close
If sObsOld <> "" Then
    sObsOld = sObsOld & vbCrLf
End If

Sql = "UPDATE PROCESSOGTI SET OBSANEXO='" & sObsOld & sObs & vbCrLf & "' where ano=" & nAno & " and numero=" & nNumproc
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdSair_Click()
CodProcesso = 0
Unload Me
End Sub

Private Sub Form_Activate()
Dim x As Integer, bAchou As Boolean, sObs As String, RdoAux2 As rdoResultset, sObsOld As String
frmProcesso.Enabled = False
If CodProcesso > 0 Then
    If AnoProcesso = lblAno.Caption And CodProcesso = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) Then
        MsgBox "Anexo não pode ser igual ao processo anexado.", vbCritical, "Atenção"
    Else
        With grdAnexo
            bAchou = False
            For x = 1 To .Rows - 1
                If .TextMatrix(x, 0) = AnoProcesso And .TextMatrix(x, 1) = CodProcesso Then
                    bAchou = True
                    Exit For
                End If
            Next
        End With
        If bAchou Then
            MsgBox "Anexo já incluso neste processo.", vbExclamation, "Atenção"
        Else
            Sql = "SELECT  ASSUNTO.NOME, PROCESSOGTI.DATAENTRADA, PROCESSOGTI.DATACANCEL, PROCESSOGTI.DATAARQUIVA, PROCESSOGTI.DATASUSPENSO "
            Sql = Sql & "FROM ASSUNTO INNER JOIN PROCESSOGTI ON ASSUNTO.CODIGO = PROCESSOGTI.CODASSUNTO "
            Sql = Sql & "Where PROCESSOGTI.ANO = " & AnoProcesso & " And PROCESSOGTI.Numero = " & CodProcesso
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                sDataEntrada = Format(!DATAENTRADA, "dd/mm/yyyy")
                If IsNull(!DataCancel) Then
                    sDataCancel = ""
                Else
                    sDataCancel = Format(!DataCancel, "dd/mm/yyyy")
                End If
                If IsNull(!DATAARQUIVA) Then
                    sDataArquiva = ""
                Else
                    sDataArquiva = Format(!DATAARQUIVA, "dd/mm/yyyy")
                End If
                If IsNull(!DATASUSPENSO) Then
                    sDataSuspenso = ""
                Else
                    sDataSuspenso = Format(!DATASUSPENSO, "dd/mm/yyyy")
                End If
            
                grdAnexo.AddItem AnoProcesso & Chr(9) & CodProcesso & Chr(9) & SubNull(!Nome) & Chr(9) & sDataEntrada & Chr(9) & sDataCancel & Chr(9) & sDataSuspenso & Chr(9) & sDataArquiva
               .Close
            
                sObs = "O processo anexado nº " & CodProcesso & "-" & RetornaDVProcesso(CodProcesso) & "/" & AnoProcesso & " foi anexado ao processo em " & Format(Now, "dd/mm/yyyy") & " por " & RetornaUsuarioFullName2(NomeDeLogin) & "."
                'sObs = "O processo anexado nº " & CodProcesso & "-" & RetornaDVProcesso(CodProcesso) & "/" & AnoProcesso & " foi anexado ao processo em " & Format(Now, "dd/mm/yyyy") & " por PROTOCOLO/ARQUIVO."
                If txtObs.Text <> "" Then txtObs.Text = txtObs.Text & vbCrLf
                txtObs.Text = txtObs.Text & sObs
                
                Sql = "SELECT * FROM PROCESSOGTI Where ANO = " & Val(lblAno.Caption) & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                sObsOld = SubNull(RdoAux2!OBSANEXO)
                RdoAux2.Close
                Sql = "UPDATE PROCESSOGTI SET OBSANEXO='" & sObsOld & sObs & vbCrLf & "' where ano=" & Val(lblAno.Caption) & " and numero=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
                cn.Execute Sql, rdExecDirect
            
            End With
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Centraliza Me
sRet = RetEventUserForm(Me.Name)
'FormHagana
lblNumProc.Caption = frmProcesso.lblNumProc.Caption
lblAno.Caption = frmProcesso.lblAno.Caption
lblAssunto.Caption = frmProcesso.cmbAssunto.Text
lblRequerente.Caption = frmProcesso.lblNomeCid.Caption
CarregaAnexo
End Sub

Private Sub CarregaAnexo()

'CARREGA TODOS OS ANEXOS
grdAnexo.Rows = 1
Sql = "SELECT ANEXO.ANOANEXO, ANEXO.NUMEROANEXO, ASSUNTO.NOME, PROCESSOGTI.DATAENTRADA, PROCESSOGTI.DATACANCEL,"
Sql = Sql & "PROCESSOGTI.DATAARQUIVA , PROCESSOGTI.DATASUSPENSO FROM  ASSUNTO INNER JOIN "
Sql = Sql & "PROCESSOGTI ON ASSUNTO.CODIGO = PROCESSOGTI.CODASSUNTO INNER JOIN  ANEXO ON PROCESSOGTI.ANO = ANEXO.ANOANEXO AND PROCESSOGTI.NUMERO = ANEXO.NUMEROANEXO "
Sql = Sql & "Where ANEXO.ANO = " & lblAno.Caption & " And ANEXO.Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sDataEntrada = Format(!DATAENTRADA, "dd/mm/yyyy")
        If IsNull(!DataCancel) Then
            sDataCancel = ""
        Else
            sDataCancel = Format(!DataCancel, "dd/mm/yyyy")
        End If
        If IsNull(!DATAARQUIVA) Then
            sDataArquiva = ""
        Else
            sDataArquiva = Format(!DATAARQUIVA, "dd/mm/yyyy")
        End If
        If IsNull(!DATASUSPENSO) Then
            sDataSuspenso = ""
        Else
            sDataSuspenso = Format(!DATASUSPENSO, "dd/mm/yyyy")
        End If
    
        grdAnexo.AddItem !ANOANEXO & Chr(9) & !NUMEROANEXO & Chr(9) & SubNull(!Nome) & Chr(9) & sDataEntrada & Chr(9) & sDataCancel & Chr(9) & sDataSuspenso & Chr(9) & sDataArquiva
        
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT * FROM PROCESSOGTI Where ANO = " & Val(lblAno.Caption) & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
txtObs.Text = SubNull(RdoAux!OBSANEXO)
RdoAux.Close


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim nNumproc As Long, nAno As Integer, aProc() As Processo, x As Integer, y As Integer

ReDim aProc(1)
nAno = Val(lblAno.Caption)
nNumproc = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
aProc(1).nAno = nAno: aProc(1).nNumproc = nNumproc
With grdAnexo
    For x = 1 To .Rows - 1
        nAno = Val(.TextMatrix(x, 0))
        nNumproc = Val(.TextMatrix(x, 1))
        ReDim Preserve aProc(UBound(aProc) + 1)
        aProc(UBound(aProc)).nAno = nAno: aProc(UBound(aProc)).nNumproc = nNumproc
    Next
End With


On Error Resume Next
For x = 1 To UBound(aProc)
    For y = 1 To UBound(aProc)
        If x <> y Then
            Sql = "INSERT ANEXO(ANO,NUMERO,ANOANEXO,NUMEROANEXO) VALUES("
            Sql = Sql & aProc(x).nAno & "," & aProc(x).nNumproc & "," & aProc(y).nAno & "," & aProc(y).nNumproc & ")"
            cn.Execute Sql, rdExecDirect
        End If
    Next
Next
On Error GoTo 0

frmProcesso.Enabled = True
If grdAnexo.Rows > 1 Then
    frmProcesso.lblAnexo.Caption = grdAnexo.Rows - 1 & " Anexo(s)"
Else
    frmProcesso.lblAnexo.Caption = "Nenhum"
End If

End Sub


Private Sub FormHagana()

'DESATIVADO
evEsp = 11

bEsp = False
If InStr(1, sRet, Format(evEsp, "000"), vbBinaryCompare) > 0 Then bEsp = True

If bEsp Then
    cmdInserir.Enabled = True
    cmdRemover.Enabled = True
Else
    cmdInserir.Enabled = False
    cmdRemover.Enabled = False
End If

End Sub

