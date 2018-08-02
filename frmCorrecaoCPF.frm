VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCorrecaoCPF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correção de CPF"
   ClientHeight    =   1155
   ClientLeft      =   10155
   ClientTop       =   5940
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5760
   Begin VB.CommandButton cmdExec 
      Caption         =   "Gerar"
      Height          =   330
      Left            =   4260
      TabIndex        =   0
      Top             =   720
      Width           =   1170
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   165
      Left            =   270
      TabIndex        =   1
      Top             =   810
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   765
      Left            =   300
      TabIndex        =   4
      Top             =   1980
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   1349
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "antigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CPFantigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CNPJantigo"
         Object.Width           =   2542
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "novo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "cpf"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "cnpj"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "endereço"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "numero"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "bairro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "quadra"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "lote"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Gerar planilha com os dados dos proprietários para correção."
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   5145
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3570
      TabIndex        =   2
      Top             =   810
      Width           =   480
   End
End
Attribute VB_Name = "frmCorrecaoCPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExec_Click()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nMes As Integer, nDia As Integer
Dim X As Integer, sArq As String, nPos As Long, nTot As Long, sCPF As String, sCNPJ As String, nCod As Long, itmX As ListItem, z As Long
'GoTo parte2
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)
X = 0
cmdExec.Enabled = False
Ocupado
Sql = "delete from codigostmp"
cn.Execute Sql, rdExecDirect

Sql = "SELECT  * FROM VWFULLIMOVEL where vwFULLIMOVEL.Ativo = 'S' order by codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sCPF = ""
        sCNPJ = ""
        nCod = 0
        If !CodCidadao < 500000 Then
            Sql = "select codcidadao,nomecidadao,cpf,cnpj from cidadao where codcidadao>500000 and nomecidadao='" & Mask(!nomecidadao) & "'"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                sCPF = SubNull(RdoAux2!CPF)
                sCNPJ = SubNull(RdoAux2!Cnpj)
                nCod = RdoAux2!CodCidadao
            End If
        End If
        Sql = "insert codigostmp (imovel,antigo,cpfantigo,cnpjantigo,novo,nome,cpf,cnpj,endereco,numero,bairro,quadra,lote) values(" & !CODREDUZIDO & "," & SubNull(!CodCidadao) & ",'" & SubNull(!CPF) & "','" & SubNull(!Cnpj) & "'," & nCod & ",'"
        Sql = Sql & Mask(!nomecidadao) & "','" & sCPF & "','" & sCNPJ & "','" & Mask(!Logradouro) & "'," & !Li_Num & ",'" & !DescBairro & "','" & Mask(Left(SubNull(!Li_Quadras), 15)) & "','" & Mask(Left(SubNull(!Li_Lotes), 15)) & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
    
   .Close
End With
parte2:
Sql = "SELECT codigostmp.imovel, codigostmp.antigo, codigostmp.cpfantigo, codigostmp.cnpjantigo, codigostmp.novo, codigostmp.endereco, codigostmp.numero, "
Sql = Sql & "codigostmp.nome , codigostmp.CPF, codigostmp.Cnpj, codigostmp.Bairro, quadra,lote FROM codigostmp INNER JOIN laseriptu ON codigostmp.imovel = laseriptu.codreduzido "
Sql = Sql & "WHERE (laseriptu.ano = 2016)"
'Sql = "select * from codigostmp order by imovel"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, "C" & CStr(!Imovel), !Imovel)
        itmX.SubItems(1) = !antigo
        itmX.SubItems(2) = !CPFantigo
        itmX.SubItems(3) = !Cnpjantigo
        itmX.SubItems(4) = !Novo
        itmX.SubItems(5) = !Nome
        itmX.SubItems(6) = !CPF
        itmX.SubItems(7) = !Cnpj
        itmX.SubItems(8) = !Endereco
        itmX.SubItems(9) = !Numero
        itmX.SubItems(10) = !Bairro
        itmX.SubItems(11) = !Quadra
        itmX.SubItems(12) = !Lote
       .MoveNext
    Loop
   .Close
End With


Liberado
cmdExec.Enabled = True
Pb.value = 0
lblPB.Caption = "0 %"
PrintExcel

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2)

'Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
Centraliza Me
Pb.value = 0
lblPB.Caption = "0 %"
End Sub

Private Sub PrintExcel()

If lvMain.ListItems.Count = 0 Then Exit Sub

Dim X As Long, Y As Long, ax As String, Scr_hdc As Long, z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For Y = 1 To lvMain.ColumnHeaders.Count
    ax = ax & RemoveSpace(lvMain.ColumnHeaders(Y).Text) & " char(255), "
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "[Table1$]", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For X = 1 To lvMain.ListItems.Count
    Rs.AddNew
    nCont = 0
    Rs.Fields(nCont).value = lvMain.ListItems(X).Text
    nCont = nCont + 1
    For Y = 2 To lvMain.ColumnHeaders.Count
         
         Rs.Fields(nCont).value = lvMain.ListItems(X).SubItems(Y - 1)
         nCont = nCont + 1
    
        
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)


End Sub

