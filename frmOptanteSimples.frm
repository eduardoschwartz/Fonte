VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmOptanteSimples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de arquivos do Simples Nacional e MEI"
   ClientHeight    =   1425
   ClientLeft      =   5985
   ClientTop       =   6480
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtArq 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   5685
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Arquivo"
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
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmOptanteSimples.frx":0000
      PICN            =   "frmOptanteSimples.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   345
      Left            =   5790
      TabIndex        =   2
      ToolTipText     =   "Executar a operação selecionada"
      Top             =   870
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Executar"
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
      MICON           =   "frmOptanteSimples.frx":00D7
      PICN            =   "frmOptanteSimples.frx":00F3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   180
      Left            =   2835
      TabIndex        =   5
      Top             =   1020
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   5205
      TabIndex        =   7
      Top             =   1005
      Width           =   720
   End
   Begin VB.Label lblPF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   2250
      TabIndex        =   6
      Top             =   1005
      Width           =   390
   End
   Begin VB.Label lblTipo 
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2850
      TabIndex        =   4
      Top             =   690
      Width           =   2685
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo selecionado.:"
      Height          =   225
      Left            =   1440
      TabIndex        =   3
      Top             =   690
      Width           =   1365
   End
End
Attribute VB_Name = "frmOptanteSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tSimples
    Codigo As Long
    Data_Inicio As String
    Data_Final As String
    Cnpj_Base As String
End Type

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPF.Caption = FormatNumber(Pb.value, 2)

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub cmdArq_Click()
Dim fName As String, cc As cCommonDlg, ff As Long, sReg As String
ff = FreeFile
Set cc = New cCommonDlg
lblTipo.Caption = ""
If cc.VBGetOpenFileName(fName, "", True, False, False, False, "Texto[*.txt]", , sPathBin, "Selecione um arquivo so Simples ou MEI", , , , False) Then
    txtArq.Text = fName
    Open fName For Input As #ff
    Do While Not EOF(1)
        Line Input #ff, sReg
        Exit Do
    Loop
    Close #ff
    If Len(sReg) = 34 Then
        lblTipo.Caption = "01-Arquivo do Simples Nacional"
    ElseIf Len(sReg) = 43 Then
        lblTipo.Caption = "02-Arquivo do MEI"
    Else
        lblTipo.Caption = "03-Arquivo inválido"
    End If
End If

End Sub

Private Sub cmdExec_Click()
Dim fName As String, nTipo As Integer, sReg As String, sCNPJ_Base As String, sData_Ini As String, sData_Fim As String
Dim aSimples() As tSimples, Sql As String, RdoAux As rdoResultset, Item As Long, z As Integer, bFind As Boolean, nCodReduz As Long
Dim nPos As Long, nTot As Long


fName = txtArq.Text
nTipo = Val(Left(lblTipo.Caption, 2))
nTot = 0
nPos = 1

If nTipo = 3 Then
    MsgBox "Arquivo inválido!", vbCritical, "Erro"
    Exit Sub
End If

'Importação do Simples Nacional
If nTipo = 1 Then
    
    Open fName For Input As #15
    Do While Not EOF(15)
        Line Input #15, sReg
        nTot = nTot + 1
    Loop
    Close #15

    ReDim aSimples(0)
    Sql = "select * from optante_simples order by cnpj_base"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Item = UBound(aSimples) + 1
            ReDim Preserve aSimples(Item)
            aSimples(Item).Cnpj_Base = !Cnpj_Base
            aSimples(Item).Codigo = !Codigo
            aSimples(Item).Data_Inicio = Format(!Data_Inicio, "dd/mm/yyyy")
            aSimples(Item).Data_Final = !Data_Final
           .MoveNext
        Loop
       .Close
    End With

    Open fName For Input As #16
    nPos = 1
    Do While Not EOF(16)
        If nPos Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
    
        Line Input #16, sReg
        sCNPJ_Base = Left(sReg, 8)
        sData_Ini = Mid(sReg, 15, 2) & "/" & Mid(sReg, 13, 2) & "/" & Mid(sReg, 9, 4)
        sData_Fim = Mid(sReg, 23, 2) & "/" & Mid(sReg, 21, 2) & "/" & Mid(sReg, 17, 4)
        If sData_Fim = "00/00/0000" Then sData_Fim = ""
        
        'procura linha do arquivo na matriz da tabela
        bFind = False
        For z = 1 To UBound(aSimples)
            With aSimples(z)
                If .Cnpj_Base = sCNPJ_Base And .Data_Inicio = sData_Ini And .Data_Final = sData_Fim Then
                    bFind = True
                    Exit For
                End If
            End With
        Next
        
        nCodReduz = nPos
        If Not bFind Then
            Sql = "select codigomob from mobiliario where SUBSTRING(cnpj, 1, 8) = '" & sCNPJ_Base & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                nCodReduz = RdoAux!codigomob
                On Error GoTo Erro
                Sql = "insert optante_simples(codigo,data_inicio,data_final,cnpj_base,timestamp) values(" & nCodReduz & ",'"
                Sql = Sql & Format(sData_Ini, "mm/dd/yyyy") & "','" & sData_Fim & "','" & sCNPJ_Base & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
                cn.Execute Sql, rdExecDirect
                On Error GoTo 0
            End If
        End If
        nPos = nPos + 1
    Loop
    Close #16
    GoTo fim
ElseIf nTipo = 2 Then
    
    Open fName For Input As #15
    Do While Not EOF(15)
        Line Input #15, sReg
        nTot = nTot + 1
    Loop
    Close #15

    ReDim aSimples(0)
    Sql = "select * from periodomei order by codigo"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Item = UBound(aSimples) + 1
            ReDim Preserve aSimples(Item)
            aSimples(Item).Cnpj_Base = !Cnpj_Base
            aSimples(Item).Codigo = !Codigo
            aSimples(Item).Data_Inicio = Format(!DataInicio, "dd/mm/yyyy")
            If Not IsNull(!Datafim) Then
                aSimples(Item).Data_Final = !Datafim
            End If
           .MoveNext
        Loop
       .Close
    End With

    Open fName For Input As #16
    nPos = 1
    Do While Not EOF(16)
        If nPos Mod 50 = 0 Then
           CallPb nPos, nTot
        End If
    
        Line Input #16, sReg
        sCNPJ_Base = Left(sReg, 8)
        sData_Ini = Mid(sReg, 15, 2) & "/" & Mid(sReg, 13, 2) & "/" & Mid(sReg, 9, 4)
        sData_Fim = Mid(sReg, 23, 2) & "/" & Mid(sReg, 21, 2) & "/" & Mid(sReg, 17, 4)
        If sData_Fim = "00/00/0000" Then sData_Fim = ""
        
        'procura linha do arquivo na matriz da tabela
        bFind = False
        For z = 1 To UBound(aSimples)
            With aSimples(z)
                If .Cnpj_Base = sCNPJ_Base And .Data_Inicio = sData_Ini And .Data_Final = sData_Fim Then
                    bFind = True
                    Exit For
                End If
            End With
        Next
        
        nCodReduz = nPos
        If Not bFind Then
            Sql = "select codigomob from mobiliario where SUBSTRING(cnpj, 1, 8) = '" & sCNPJ_Base & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                nCodReduz = RdoAux!codigomob
                On Error GoTo Erro
                If sDataFim <> "" Then
                    Sql = "insert periodomei(codigo,datainicio,datafim,cnpj_base) values(" & nCodReduz & ",'"
                    Sql = Sql & Format(sData_Ini, "mm/dd/yyyy") & "','" & Format(sData_Fim, "mm/dd/yyyy") & "','" & sCNPJ_Base & "')"
                Else
                    Sql = "insert periodomei(codigo,datainicio,cnpj_base) values(" & nCodReduz & ",'"
                    Sql = Sql & Format(sData_Ini, "mm/dd/yyyy") & "','" & sCNPJ_Base & "')"
                End If
                cn.Execute Sql, rdExecDirect
                
                On Error GoTo 0
            End If
        End If
        nPos = nPos + 1
    Loop
    Close #16
    GoTo fim
End If

Exit Sub
Erro:
If rdoErrors(1).Number = 2627 Then
    Resume Next
Else
    MsgBox rdoErrors(1).Description
    
End If

Exit Sub
fim:
lblTipo.Caption = ""
txtArq.Text = ""
Pb.value = 0
lblPF.Caption = "0"
lblTot.Caption = "0"
MsgBox "Importação finalizada.", vbInformation, "Atenção"


End Sub

