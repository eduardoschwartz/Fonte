VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmConversor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversor do Sistema"
   ClientHeight    =   2385
   ClientLeft      =   6480
   ClientTop       =   4755
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5265
   Begin Tributacao.XP_ProgressBar Pb 
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   585
      Width           =   3300
      _ExtentX        =   5821
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
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   1065
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   5025
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Emitir Relatório"
      Top             =   570
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConversor.frx":0000
      PICN            =   "frmConversor.frx":001C
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
Attribute VB_Name = "frmConversor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

Private Sub cmdExec_Click()
Dim nPos As Long, nTotal As Long, nMax As Long
Pb.SetFocus
txtLog.text = ""
Sql = "SELECT * FROM TRIBLOCAL..MOBILIARIO ORDER BY CODIGOMOB"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotal = .RowCount
    Do Until .EOF
        nPos = .AbsolutePosition
        If nPos Mod 10 = 0 Then
           CallPb nPos, CLng(nTotal)
        End If
        Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODIGOMOB=" & !CODIGOMOB
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount = 0 Then
            Sql = "INSERT MOBILIARIO SELECT * FROM TRIBLOCAL..MOBILIARIO WHERE CODIGOMOB=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT MOBILIARIOATIVIDADEISS SELECT * FROM TRIBLOCAL..MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT MOBILIARIOATIVIDADETL SELECT * FROM TRIBLOCAL..MOBILIARIOATIVIDADETL WHERE CODIGOMOB=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT MOBILIARIOATIVIDADEVS SELECT * FROM TRIBLOCAL..MOBILIARIOATIVIDADEVS WHERE CODMOBILIARIO=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT MOBILIARIOENDENTREGA SELECT * FROM TRIBLOCAL..MOBILIARIOENDENTREGA WHERE CODMOBILIARIO=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT MOBILIARIOEVENTO SELECT * FROM TRIBLOCAL..MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT MOBILIARIOHIST SELECT * FROM TRIBLOCAL..MOBILIARIOHIST WHERE CODMOBILIARIO=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT MOBILIARIOPROPRIETARIO SELECT * FROM TRIBLOCAL..MOBILIARIOPROPRIETARIO WHERE CODMOBILIARIO=" & !CODIGOMOB
            cn.Execute Sql, rdExecDirect
            RdoAux2.Close
            
            Sql = "SELECT CODCIDADAO FROM TRIBLOCAL..MOBILIARIOPROPRIETARIO WHERE CODMOBILIARIO=" & !CODIGOMOB
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                'rdoAux2.Close
                Do Until RdoAux2.EOF
                    Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO"
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    nMax = RdoAux3!MAXIMO + 1
                    RdoAux3.Close
                    Sql = "INSERT CIDADAO SELECT " & nMax & ",nomecidadao,codnew,excluir,cpf,cnpj,codlogradouro,numimovel,complemento,codbairro,codcidade,"
                    Sql = Sql & "siglauf,cep,telefone,email,rg,nomelogradouro,orgao,nomecidade,nomebairro,nomeuf FROM TRIBLOCAL..CIDADAO WHERE CODCIDADAO=" & RdoAux2!CodCidadao
                    cn.Execute Sql, rdExecDirect
                    RdoAux2.MoveNext
                Loop
            End If
                    
            txtLog = txtLog & !CODIGOMOB & " - " & !RAZAOSOCIAL & vbCrLf
        End If
       .MoveNext
    Loop
   .Close
End With

MsgBox "Conversão Finalizada !!!", vbInformation, "Atenção"

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.Value = (nPosF * 100) / nTotal
Else
   Pb.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

