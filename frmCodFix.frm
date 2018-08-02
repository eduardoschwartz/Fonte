VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCodFix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correção de códigos"
   ClientHeight    =   4785
   ClientLeft      =   5415
   ClientTop       =   3615
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   4005
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H8000000D&
      Caption         =   "Parar"
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Top             =   630
      Width           =   915
   End
   Begin VB.CommandButton cmdFixList 
      BackColor       =   &H8000000D&
      Caption         =   "Corrigir"
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   270
      Width           =   915
   End
   Begin VB.ListBox lstCNPJ 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   90
      TabIndex        =   13
      Top             =   135
      Width           =   2670
   End
   Begin VB.CommandButton cmdFix 
      BackColor       =   &H8000000D&
      Caption         =   "Corrigir"
      Height          =   285
      Left            =   2430
      TabIndex        =   12
      Top             =   3015
      Width           =   1320
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   150
      Left            =   45
      TabIndex        =   9
      Top             =   4545
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtCNPJ 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1125
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1125
      Width           =   2760
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1845
      Width           =   3840
   End
   Begin VB.CommandButton cmdCNPJ 
      Caption         =   "Carregar"
      Height          =   285
      Left            =   3105
      TabIndex        =   3
      Top             =   1485
      Width           =   780
   End
   Begin MSComctlLib.ListView lvCod 
      Height          =   2190
      Left            =   45
      TabIndex        =   4
      Top             =   2250
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Códigos"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Débito"
         Object.Width           =   1235
      EndProperty
   End
   Begin esMaskEdit.esMaskedEdit mskCNPJ 
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   1485
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   503
      MouseIcon       =   "frmCodFix.frx":0000
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
      MaxLength       =   18
      Mask            =   "99.999.999/9999-99"
      SelText         =   ""
      Text            =   "__.___.___/____-__"
      HideSelection   =   -1  'True
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "0 de 0"
      Height          =   195
      Left            =   2385
      TabIndex        =   16
      Top             =   4185
      Width           =   1410
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3105
      TabIndex        =   11
      Top             =   2565
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Códigos......:"
      Height          =   225
      Index           =   2
      Left            =   2340
      TabIndex        =   10
      Top             =   2565
      Width           =   930
   End
   Begin VB.Label lblCod 
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3285
      TabIndex        =   8
      Top             =   2295
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cód Master.:"
      Height          =   225
      Index           =   1
      Left            =   2340
      TabIndex        =   7
      Top             =   2295
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ Novo.:"
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   1530
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ velho.:"
      Height          =   225
      Index           =   11
      Left            =   90
      TabIndex        =   2
      Top             =   1170
      Width           =   1020
   End
End
Attribute VB_Name = "frmCodFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bStop As Boolean

Private Sub cmdCNPJ_Click()
Dim sCNPJ As String, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim itmX As ListItem, z As Long, sNome As String, nCodReduz As Long, sDebito As String
Dim nTotal As Long, nPos As Long

z = SendMessage(lvCod.hwnd, LVM_DELETEALLITEMS, 0, 0)
txtNome.Text = ""
lblCod.Caption = "000000"
lblTotal.Caption = "0000"
sCNPJ = txtCNPJ.Text
nPos = 1

If mskCNPJ.ClipText = "" Then
    MsgBox "Digite um CNPJ.", vbExclamation, "Atenção"
    mskCNPJ.SetFocus
    Exit Sub
End If

If Not ValidaCGC(mskCNPJ.ClipText) Then
'    MsgBox "CNPJ inválido!", vbExclamation, "Atenção"
    Exit Sub
End If

Ocupado
Sql = "select codcidadao,nomecidadao from cidadao where cnpj='" & sCNPJ & "' or cnpj='" & mskCNPJ.ClipText & "' or cnpj='" & lstCNPJ.Text & "' order by codcidadao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotal = .RowCount
    lblTotal.Caption = Format(nTotal, "0000")
    If nTotal > 0 Then
        txtNome.Text = !nomecidadao
        lblCod.Caption = !CodCidadao
    End If
    
    Do Until .EOF
        nCodReduz = !CodCidadao
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTotal
        End If
        Sql = "select anoexercicio from debitoparcela where codreduzido=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sDebito = "S"
        Else
            sDebito = "N"
        End If
        RdoAux2.Close
        
        Set itmX = lvCod.ListItems.Add(, , Format(nCodReduz, "000000"))
        itmX.SubItems(1) = sDebito
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
Liberado
pBar.Value = 0

End Sub

Private Sub cmdFix_Click()
Dim sCNPJ As String, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim itmX As ListItem, z As Long, sNome As String, nCodReduz As Long, sDebito As String
Dim nTotal As Long, nPos As Long, nCodMaster As Long, bDebito As Boolean, nSeqNew As Integer
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
'Dim nStatusLanc As Integer, sDataVencto As String, sDataBase As String, nNumLivro As Integer, nPagLivro As Integer
'Dim nNumCertidao As Integer, sDataIncricao As String, sDataAjuiza As String, sUsuario As String

If lvCod.ListItems.Count < 2 Then
    MsgBox "Nada a corrigir.", vbExclamation, "Atenção"
    Exit Sub
End If

nCodMaster = CLng(lblCod.Caption)
nTotal = lvCod.ListItems.Count
nPos = 1

Ocupado

For Each itmX In lvCod.ListItems
    If nPos Mod 10 = 0 Then
        CallPb nPos, nTotal
    End If
    bDebito = False
    If itmX.Text <> lblCod.Caption Then
        nCodReduz = CLng(itmX.Text)
        If itmX.ListSubItems(1) = "S" Then
            bDebito = True
        End If
        
        'se tiver debito transfere
        If bDebito Then
            Sql = "select * from debitoparcela where codreduzido=" & nCodReduz & " order by anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                Do Until .EOF
                    DoEvents
                    nAno = !AnoExercicio
                    nLanc = !CodLancamento
                    nSeq = !SeqLancamento
                    nParc = !NumParcela
                    nCompl = !CODCOMPLEMENTO
                   
                   'busca maior seq da codmaster para aquele ano,lanc
                    Sql = "select max(seqlancamento) as maxseq from debitoparcela where codreduzido=" & nCodMaster & " and "
                    Sql = Sql & "anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If IsNull(RdoAux2!maxseq) Then
                        nSeqNew = 0
                    Else
                        nSeqNew = RdoAux2!maxseq + 1
                    End If
                               
                    'atualiza o lancamento velho no codigo master
                    'debito parcela
                    Sql = "update debitoparcela set codreduzido=" & nCodMaster & ",seqlancamento=" & nSeqNew & " where codreduzido=" & nCodReduz & " and "
                    Sql = Sql & "anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    cn.Execute Sql, rdExecDirect
                                       
                    'debito tributo
                    Sql = "update debitotributo set codreduzido=" & nCodMaster & ",seqlancamento=" & nSeqNew & " where codreduzido=" & nCodReduz & " and "
                    Sql = Sql & "anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    cn.Execute Sql, rdExecDirect
                   
                    'parceladocumento
                    Sql = "update parceladocumento set codreduzido=" & nCodMaster & ",seqlancamento=" & nSeqNew & " where codreduzido=" & nCodReduz & " and "
                    Sql = Sql & "anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    cn.Execute Sql, rdExecDirect
                   
                    'debito pago
                    Sql = "update debitopago set codreduzido=" & nCodMaster & ",seqlancamento=" & nSeqNew & " where codreduzido=" & nCodReduz & " and "
                    Sql = Sql & "anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    cn.Execute Sql, rdExecDirect
                   
                    'obs parcela
                    Sql = "update obsparcela set codreduzido=" & nCodMaster & ",seqlancamento=" & nSeqNew & " where codreduzido=" & nCodReduz & " and "
                    Sql = Sql & "anoexercicio=" & nAno & " and codlancamento=" & nLanc & " and seqlancamento=" & nSeq & " and numparcela=" & nParc & " and codcomplemento=" & nCompl
                    cn.Execute Sql, rdExecDirect
                    
                   .MoveNext
                Loop
               .Close
            End With
            
        End If
        
       'atualiza o cnpj para 14 posições
        Sql = "update cidadao set cnpj='" & mskCNPJ.ClipText & "' where codcidadao=" & nCodMaster
        cn.Execute Sql, rdExecDirect
        
       'grava o codigo cidadao na tabela bkp
        Sql = "insert fixcodiss(codold,codnew) values(" & nCodReduz & "," & nCodMaster & ")"
        cn.Execute Sql, rdExecDirect
                
       'apaga o codigo do cidadao
        Sql = "delete from cidadao where codcidadao=" & nCodReduz
        cn.Execute Sql, rdExecDirect
        
    End If
    nPos = nPos + 1
Next

Liberado
cmdCNPJ_Click
'MsgBox "Fim"

End Sub

Private Sub cmdFixList_Click()
Dim x As Integer, nTotal As Long

bStop = False
nTotal = lstCNPJ.ListCount - 1
For x = 0 To lstCNPJ.ListCount - 1
    If bStop Then Exit For
    lblStatus.Caption = sTr(x) & " de " & sTr(nTotal)
    lblStatus.Refresh
    lstCNPJ.ListIndex = x
    If lvCod.ListItems.Count > 1 Then
        cmdFix_Click
    End If
Next

End Sub

Private Sub cmdStop_Click()
bStop = True
DoEvents
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset

Centraliza Me


Sql = "select distinct cnpj from cidadao where LEN(cnpj)>=14 order by cnpj"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstCNPJ.AddItem !Cnpj
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub lstCNPJ_Click()
If lstCNPJ.ListIndex = -1 Then Exit Sub

txtCNPJ.Text = RetornaNumero(lstCNPJ.Text)
txtCNPJ_LostFocus
cmdCNPJ_Click

End Sub

Private Sub mskCNPJ_GotFocus()
mskCNPJ.SetFocus
End Sub

Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
Tweak txtCNPJ, KeyAscii, IntegerPositive
End Sub


Private Sub CallPb(nVal As Long, nTot As Long)
If ((nVal * 100) / nTot) <= 100 Then
   pBar.Value = (nVal * 100) / nTot
Else
   pBar.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Sub txtCNPJ_LostFocus()
If Len(txtCNPJ.Text) >= 14 Then
    mskCNPJ.Text = Format(Right(txtCNPJ.Text, 14), "00\.000\.000/0000-00")
End If
End Sub
