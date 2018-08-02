VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCancelDebito 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento de Débitos"
   ClientHeight    =   4620
   ClientLeft      =   4965
   ClientTop       =   3270
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8805
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   2325
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   8745
      Begin VB.CheckBox chkKeep 
         Caption         =   "Manter na tela de Extrato"
         Enabled         =   0   'False
         Height          =   240
         Left            =   5940
         TabIndex        =   15
         Top             =   270
         Width           =   2355
      End
      Begin esMaskEdit.esMaskedEdit mskDataProc 
         Height          =   300
         Left            =   6210
         TabIndex        =   3
         Top             =   690
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         BackColor       =   15658734
         MouseIcon       =   "frmCancelDebito.frx":0000
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
         BorderStyle     =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
         Locked          =   -1  'True
      End
      Begin VB.TextBox txtProc 
         BackColor       =   &H00EEEEEE&
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   645
         Width           =   1470
      End
      Begin VB.TextBox txtMotivo 
         Height          =   1140
         Left            =   2205
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   4950
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmCancelDebito.frx":001C
         Left            =   2205
         List            =   "frmCancelDebito.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3315
      End
      Begin prjChameleon.chameleonButton cmdHelp 
         Height          =   345
         Left            =   7290
         TabIndex        =   12
         ToolTipText     =   "Ajuda desta Tela"
         Top             =   1440
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Ajuda"
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
         MICON           =   "frmCancelDebito.frx":010A
         PICN            =   "frmCancelDebito.frx":0126
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
         Left            =   7290
         TabIndex        =   13
         ToolTipText     =   "Sair da Tela"
         Top             =   1845
         Width           =   1365
         _ExtentX        =   2408
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCancelDebito.frx":0280
         PICN            =   "frmCancelDebito.frx":029C
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
         Height          =   345
         Left            =   7290
         TabIndex        =   14
         ToolTipText     =   "Executar o Cancelamento das Parcelas selecionadas"
         Top             =   1035
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Executar"
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
         MICON           =   "frmCancelDebito.frx":030A
         PICN            =   "frmCancelDebito.frx":0326
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
         Caption         =   "Data do Processo....:"
         Height          =   225
         Index           =   3
         Left            =   4620
         TabIndex        =   9
         Top             =   690
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo sem DV:"
         Height          =   225
         Index           =   2
         Left            =   165
         TabIndex        =   8
         Top             =   705
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo Cancelamento.......:"
         Height          =   225
         Index           =   1
         Left            =   165
         TabIndex        =   7
         Top             =   1110
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cancelamento......:"
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   285
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   2190
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3863
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmCancelDebito.frx":0480
   End
   Begin VB.Label lblsupervisor 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7410
      TabIndex        =   11
      Top             =   2970
      Width           =   1245
   End
   Begin VB.Label lblSup 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   7410
      TabIndex        =   10
      Top             =   2490
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmCancelDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset
Dim sRet As String
Dim evDel As Integer, evSus As Integer, evReat As Integer
Dim bDel As Boolean

Private Sub cmbTipo_Click()

If cmbTipo.ListIndex = 0 Then
    chkKeep.Enabled = True
Else
    chkKeep.Enabled = False
End If

Select Case cmbTipo.ListIndex
    Case 0, 4
        txtProc.Locked = True
        txtProc.BackColor = Kde
    Case 1, 2, 3, 6, 7
        txtProc.Locked = False
        txtProc.BackColor = Branco
    Case Else
        Exit Sub
End Select

End Sub

Private Sub cmdCancel_Click()
Dim x As Integer
Dim nStatus As Integer
Dim nCodReduz As Long
Dim nAno As Integer
Dim nLanc As Integer
Dim nSeq As Integer
Dim nParc As Integer
Dim nCompl As Integer
Dim Achou As Boolean, sDebito As String
Dim bSupervisor As Boolean, sNomeUser As String, nLivro As Integer, nPagina As Integer
Dim qd As New rdoQuery, nVP As Double, nVJ As Double, nVM As Double, nVC As Double, nVT As Double, bAjuizado As Boolean

If grdTemp.Rows = 1 Then Exit Sub

If Not IsDate(mskDataProc.Text) And cmbTipo.ListIndex = 1 Then
    MsgBox "data invalida"
    Exit Sub
End If

evDel = 4: evSus = 17
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True
bSupervisor = bDel
    If NomeDeLogin <> "JOSIANE" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "EDUARDO" And NomeDeLogin <> "ANA" And NomeDeLogin <> "NOELI" Then
        If cmbTipo.ListIndex = 3 Or cmbTipo.ListIndex = 5 Then
            If InStr(1, sRet, Format(evSus, "000"), vbBinaryCompare) = 0 Then
                MsgBox "O Usuário " & NomeDeLogin & " não possue permissão para suspender este(s) lancamento(s).", vbCritical, "Alerta de Segurança"
                Exit Sub
            End If
        End If
End If

If cmbTipo.ListIndex = 0 Then
    If chkKeep.value = vbChecked Then
        nStatus = 8
    Else
        nStatus = 5 'cancelado por erro
    End If
    Achou = False
    
    With grdTemp
        For x = 1 To grdTemp.Rows - 1
            If CDate(grdTemp.TextMatrix(x, 10)) <> Format(Now, "dd/mm/yyyy") Then
               Achou = True
               Exit For
            End If
        Next
    End With

    If Achou And Not bSupervisor Then
        ButtonText(0) = "Supervisor"
        ButtonText(1) = "Cancelar"
        'Set up the CBT hook
        hInst = GetWindowLong(Me.hwnd, GWL_HINSTANCE)
        Thread = GetCurrentThreadId()
        hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, Thread)
        retval = MsgBox("O Usuário " & sNomeUser & " não possue permissão para cancelar parcelas vencidas." & vbCrLf & "Solicite um Supervisor ou cancele a operação", vbInformation + vbYesNo, "Alerta de Segurança")
        If retval = vbYes Then
          '  frmMonitor.show vbModal
        Else
            Exit Sub
        End If
    End If
    
    If lblSup.Caption = 0 Then
        Achou = False
        For x = 1 To grdTemp.Rows - 1
            If CDate(grdTemp.TextMatrix(x, 10)) <> Format(Now, "dd/mm/yyyy") Then
                Achou = True
                Exit For
            End If
        Next
        If Achou And Not bSupervisor Then
            ButtonText(0) = "Supervisor"
            ButtonText(1) = "Cancelar"
            'Set up the CBT hook
            hInst = GetWindowLong(Me.hwnd, GWL_HINSTANCE)
            Thread = GetCurrentThreadId()
            hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, Thread)
            retval = MsgBox("O Usuário " & sNomeUser & "  não possue permissão para cancelar parcelas com Data Base antiga." & vbCrLf & "Solicite um Supervisor ou cancele a operação", vbInformation + vbYesNo, "Alerta de Segurança")
            If retval = vbYes Then
                
                ' frmMonitor.show vbModal
            Else
                Exit Sub
            End If
        End If
   End If
ElseIf cmbTipo.ListIndex = 1 Then
    nStatus = 8 'cancelado por recurso
    If Not IsDate(mskDataProc.Text) Then
       MsgBox "Data do Processo inválido.", vbExclamation, "Atenção"
       Exit Sub
    End If
ElseIf cmbTipo.ListIndex = 6 Then
'    nStatus = 28 'compensado por recurso
    nStatus = 6 'compensado por recurso
    If Not IsDate(mskDataProc.Text) Then
       MsgBox "Data do Processo inválido.", vbExclamation, "Atenção"
       Exit Sub
    End If
ElseIf cmbTipo.ListIndex = 7 Then
'    nStatus = 28 'compensado por recurso
    nStatus = 32 'compensação tributária
    If Not IsDate(mskDataProc.Text) Then
       MsgBox "Data do Processo inválido.", vbExclamation, "Atenção"
       Exit Sub
    End If
ElseIf cmbTipo.ListIndex = 4 Then
    nStatus = 12 'cancelado por duplicidade
ElseIf cmbTipo.ListIndex = 5 Then
    nStatus = 27 'retido pelo tomador
ElseIf cmbTipo.ListIndex = 2 Then
    nStatus = 14 'cancelado sem movimento
    If Not IsDate(mskDataProc.Text) Then
       MsgBox "O cancelamento será registrado sem a Data do Processo.", vbExclamation, "Atenção"
       'Exit Sub
    End If
ElseIf cmbTipo.ListIndex = 3 Then
    nStatus = 19 'suspenso/tramite
    If Not IsDate(mskDataProc.Text) Then
       MsgBox "Data do Processo inválido.", vbExclamation, "Atenção"
       Exit Sub
    End If
End If

If Trim$(txtMotivo.Text) = "" And (cmbTipo.ListIndex = 1 Or cmbTipo.ListIndex = 2 Or cmbTipo.ListIndex = 3) Then
     MsgBox "Digite o Motivo do cancelamento/suspensão.", vbCritical, "Atenção"
     txtMotivo.SetFocus
     Exit Sub
End If

If MsgBox("Executar a operação selecionada nos débitos acima selecionados ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    nCodReduz = Val(frmDebitoImob.txtCod.Text)
        
    ConectaIntegrativa
    For x = 1 To grdTemp.Rows - 1
        With grdTemp
            nAno = Val(.TextMatrix(x, 0))
            nLanc = Val(Left$(.TextMatrix(x, 1), 3))
            nSeq = Val(.TextMatrix(x, 2))
            nParc = Val(.TextMatrix(x, 3))
            nCompl = Val(.TextMatrix(x, 4))
        End With
        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC =" & nStatus & " WHERE CODREDUZIDO=" & nCodReduz
        Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
        Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
        Sql = "DELETE FROM  DEBITOCANCEL WHERE CODREDUZIDO=" & nCodReduz
        Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
        Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
        If Trim$(txtMotivo.Text) <> "" Then
            Sql = "INSERT DEBITOCANCEL(NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,USERID,DATACANCEL,MOTIVO) VALUES('"
            Sql = Sql & txtProc.Text & "'," & nCodReduz & "," & nAno & "," & nLanc & ","
            Sql = Sql & nSeq & "," & nParc & "," & nCompl & "," & RetornaUsuarioID(NomeDeLogin) & ",'"
            Sql = Sql & Format(Now, "mm/dd/yyyy") & "','" & Mask(txtMotivo.Text) & "')"
            cn.Execute Sql, rdExecDirect
        End If
        sDebito = "Ano:" & nAno & " Código:" & nCodReduz & " Lançamento:" & nLanc
        sDebito = sDebito & " Seq:" & nSeq & " Parcela:" & nParc & " Compl:" & nCompl
        sDebito = sDebito & " Vencto:" & grdTemp.TextMatrix(x, 6) & " Supervisor: " & lblsupervisor.Caption
        sDebito = sDebito & " Motivo: " & txtMotivo.Text
        Log Form, Me.Caption, Exclusão, sDebito
        
        If cmbTipo.ListIndex = 2 Then
            Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=0 WHERE CODREDUZIDO=" & nCodReduz
            Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & 5 & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO=13"
            cn.Execute Sql, rdExecDirect
        End If
        
       '****INTEGRATIVA******
                   
        If grdTemp.TextMatrix(x, 7) = "S" Then
            nVP = 0: nVJ = 0: nVM = 0: nVC = 0: nVT = o
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = nCodReduz
            qd(2) = nAno
            qd(3) = nAno
            qd(4) = nLanc
            qd(5) = nLanc
            qd(6) = nSeq
            qd(7) = nSeq
            qd(8) = nParc
            qd(9) = nParc
            qd(10) = nCompl
            qd(11) = nCompl
            qd(12) = 1
            qd(13) = 99
            qd(14) = Format(Now, "mm/dd/yyyy")
            qd(15) = NomeDoUsuario
            Set RdoAux = qd.OpenResultset(rdOpenKeyset)
            With RdoAux
                Do Until .EOF
                    nVP = nVP + !ValorTributo
                    nVJ = nVJ + !ValorJuros
                    nVM = nVM + !ValorMulta
                    nVC = nVC + !ValorCorrecao
                    nVT = nVT + !ValorTotal
                    nLivro = Val(SubNull(!NUMLIVRO))
                    nPagina = Val(SubNull(!PAGINA))
                    bAjuizado = IIf(IsNull(!dataajuiza), 0, 1)
                   .MoveNext
                Loop
               .Close
            End With
                    
           '*** VERIFICA SE O PARCELAMENTO JÁ EXISTE NA TABELA ACORDOS **
           Sql = "select * from cancelamentos where iddevedor=" & nCodReduz & " and exercicio=" & nAno & " and lancamento=" & nLanc & "and seq=" & nSeq & "and "
           Sql = Sql & "nroparcela=" & nParc & " and complparcela=" & nCompl
           Set RdoAux = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           If RdoAux.RowCount = 0 Then
               If frmMdi.frTeste.Visible = False Then
                    'GRAVA O ACORDO
                    Sql = "insert cancelamentos(dtCancelamento, idDevedor, NroLivro, NroFolha, Seq, Lancamento, Exercicio, VlrOriginal, VlrJuros, VlrMulta, VlrTotal, nroParcela, "
                    Sql = Sql & "ComplParcela, Ajuizado, DtGeracao) values ('" & Format(Now, "mm/dd/yyyy") & "'," & nCodReduz & "," & nLivro & "," & nPagina & "," & nSeq & "," & nLanc & "," & nAno & ","
                    Sql = Sql & Virg2Ponto(CStr(nVP)) & "," & Virg2Ponto(CStr(nVJ)) & "," & Virg2Ponto(CStr(nVM)) & "," & Virg2Ponto(CStr(nVT)) & "," & nParc & "," & nCompl & "," & IIf(bAjuizado, 1, 0) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                    cnInt.Execute Sql, rdExecDirect
               End If
           End If
        
        End If
       '*********************
proximo:
    Next
    MsgBox "Os débitos foram alterados com sucesso.", vbInformation, "Atenção"
    Unload Me
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.ZOrder 0
End Sub

Private Sub Form_Load()

Centraliza Me
Me.Top = Me.Top + 1200
Ocupado
cmbTipo.ListIndex = 0
sRet = RetEventUserForm(Me.Name)
CarregaLista
Liberado

End Sub

Private Sub CarregaLista()
Dim x As Integer, nCodReduz As Long
Dim sAno As String, sLanc As String, sSeq As String, sParc As String
Dim sComp As String, sSit As String, sVencto As String, sDA As String
Dim sAj As String, nValorPrincipal As Double, sDataBase As String

With frmDebitoImob.grdExtrato
    nCodReduz = Val(frmDebitoImob.txtCod.Text)
    For x = 1 To .Rows
        If .CellText(x, 12) = "S" Then
           sAno = .CellText(x, 1)
           sLanc = .CellText(x, 2)
           sSeq = .CellText(x, 3)
           sParc = IIf(.CellText(x, 4) = "Unica", "00", .CellText(x, 4))
           sComp = .CellText(x, 5)
           sSit = Left$(.CellText(x, 6), 2)
           sVencto = .CellText(x, 7)
           sDA = .CellText(x, 8)
           sAj = .CellText(x, 9)
           nValorPrincipal = .CellText(x, 10)
           
           Sql = "SELECT CODREDUZIDO,DATADEBASE FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz
           Sql = Sql & " AND ANOEXERCICIO=" & Val(sAno) & " AND CODLANCAMENTO=" & Val(sLanc) & " AND "
           Sql = Sql & " SEQLANCAMENTO=" & Val(sSeq) & " AND NUMPARCELA=" & Val(sParc) & " AND "
           Sql = Sql & " CODCOMPLEMENTO=" & Val(sComp)
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
                If .RowCount > 0 Then
                    sDataBase = Format(!DATADEBASE, "dd/mm/yyyy")
                Else
                    Exit Sub
                End If
               .Close
           End With
           
           grdTemp.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
             sComp & Chr(9) & sSit & Chr(9) & sVencto & Chr(9) & sDA & Chr(9) & sAj & Chr(9) & _
             FormatNumber(nValorPrincipal, 2) & Chr(9) & sDataBase
             
          
        End If
    Next

End With

End Sub

Private Sub txtProc_LostFocus()
Dim sValidaProc As String

txtProc.Text = Replace(txtProc.Text, "-", "")
If Trim(txtProc.Text) = "" Then Exit Sub
sValidaProc = ValidaProcesso(txtProc.Text)
If sValidaProc <> "OK" Then
    MsgBox sValidaProc, vbCritical, "Atenção"
End If
LimpaMascara mskDataProc
mskDataProc.Text = Format(RetornaDataProcesso(ExtraiNumeroProcesso(txtProc.Text), ExtraiAnoProcesso(txtProc.Text)), "dd/mm/yyyy")
'mskDataProc.Text = Format(RetornaDataProcesso(Val(Left$(txtProc.Text, Len(txtProc.Text) - 5)), Val(Right$(txtProc.Text, 4))), "dd/mm/yyyy")
If mskDataProc.Text = "01/01/1899" Then
     LimpaMascara mskDataProc
End If

End Sub
