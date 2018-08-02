VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNovaGIA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nova GIA"
   ClientHeight    =   3540
   ClientLeft      =   16575
   ClientTop       =   5085
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   7665
   Begin Tributacao.XP_ProgressBar pBar2 
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   3240
      Width           =   3795
      _ExtentX        =   6694
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
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   2970
      Width           =   3795
      _ExtentX        =   6694
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
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdImport 
      Height          =   360
      Left            =   6255
      TabIndex        =   2
      ToolTipText     =   "Importar o arquivo para o GTI"
      Top             =   3060
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Importar"
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
      MICON           =   "frmNovaGIA.frx":0000
      PICN            =   "frmNovaGIA.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   1260
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   8388608
      DialogTitle     =   "Selecione o arquivo de GIA"
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   360
      Left            =   4860
      TabIndex        =   0
      ToolTipText     =   "Selecione o arquivo a importar"
      Top             =   3060
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Arquivo"
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
      MICON           =   "frmNovaGIA.frx":00A4
      PICN            =   "frmNovaGIA.frx":00C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2820
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4974
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Razão Social"
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "CNPJ"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Insc.Estadual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Inscrição "
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmNovaGIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As ADODB.Connection

Private Sub cmdArq_Click()
Dim Rs As ADODB.Recordset, strQuery As String
Dim itmX As ListItem, z As Long, sArq As String
Dim nCodReduz As Long, Sql As String, Rdoaux As rdoResultset

z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)
sArq = ""
cDialog.FileName = "*.mdb"
cDialog.InitDir = App.Path & "\bin"
cDialog.Filter = "Arquivo GIA (*.mdb)"
cDialog.FilterIndex = 1
cDialog.flags = cdlOFNFileMustExist

cDialog.ShowOpen
sArq = cDialog.FileName
If sArq = "*.mdb" Then Exit Sub
On Error GoTo Erro
Set cnn = New ADODB.Connection
cnn.CursorLocation = adUseClient
cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sArq & ";Persist Security Info=False;Jet OLEDB:Database Password=kamisama2"
cnn.Open
If cnn.State = 0 Then
    cnn.Close
    GoTo Erro
End If
Ocupado
Sql = "SELECT * FROM tblCONTRIBUINTE ORDER BY RAZÃOSOCIAL"
Set Rs = New ADODB.Recordset
With Rs
    Set .ActiveConnection = cnn
   .CursorType = adOpenStatic
   .Source = Sql
   .Open
    While Not .EOF
        If .AbsolutePosition Mod 5 = 0 Then
            CallPb CLng(.AbsolutePosition), CLng(.RecordCount)
        End If
        nCodReduz = 0
        Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CNPJ='" & RetornaNumero(!Cnpj) & "'"
        Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If Rdoaux.RowCount > 0 Then
            nCodReduz = Rdoaux!codigomob
        End If
        Rdoaux.Close
        
        Set itmX = lvMain.ListItems.Add(, !Cnpj, !RAZÃOSOCIAL)
        itmX.SubItems(1) = !Cnpj
        itmX.SubItems(2) = !IE
        itmX.SubItems(3) = nCodReduz
       .MoveNext
    Wend
   .Close
End With
Liberado
PBar.value = 0
PBar.Color = vbWhite

      
Exit Sub
Erro:
MsgBox "Erro na conexão do arquivo.", vbCritical, "Atenção"

End Sub

Private Sub cmdImport_Click()
Importar
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub Importar()
Dim Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, strQuery As String
Dim X As Integer, Sql As String, Rdoaux As rdoResultset
Dim sCNPJ As String, nCodReduz As Long, sIE As String, dRef As Date

If lvMain.ListItems.Count = 0 Then
    MsgBox "Nada a importar.", vbExclamation, "Atenção"
    Exit Sub
End If
Me.Refresh
If MsgBox("Importar este arquivo?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

For X = 1 To lvMain.ListItems.Count
    CallPb CLng(X), CLng(lvMain.ListItems.Count)
    sIE = lvMain.ListItems(X).SubItems(2)
    nCodReduz = lvMain.ListItems(X).SubItems(3)
    If nCodReduz = 0 Then GoTo PROXIMO2
    Sql = "SELECT * FROM tblGIA WHERE IE='" & sIE & "' ORDER BY NroGIA"
    Set Rs1 = New ADODB.Recordset
    With Rs1
        Set .ActiveConnection = cnn
       .CursorType = adOpenStatic
       .Source = Sql
       .Open
        While Not .EOF
            CallPb2 CLng(.AbsolutePosition), CLng(.RecordCount)
            dRef = !REF1
            Sql = "SELECT * FROM tblDetalhesCFOPs WHERE NroGIA=" & !nroGIA & " ORDER BY CFOP"
            Set Rs2 = New ADODB.Recordset
            With Rs2
                Set .ActiveConnection = cnn
               .CursorType = adOpenStatic
               .Source = Sql
               .Open
                While Not .EOF
                    If RetornaCFOP(!CFOP) Then
                        Sql = "SELECT * FROM GIADETALHE WHERE CODREDUZIDO=" & nCodReduz & " AND REF='" & Format(dRef, "mm/dd/yyyy") & "' AND "
                        Sql = Sql & "NUMGIA=" & !nroGIA & " AND CFOP=" & RetornaNumero(!CFOP)
                        Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If Rdoaux.RowCount > 0 Then
                            Rdoaux.Close
                            GoTo proximo
                        End If
                        Rdoaux.Close
                        
                        Sql = "INSERT GIADETALHE(CODREDUZIDO,REF,NUMGIA,CFOP,BASECALCULO,ISENTASNTRIB,OUTRAS) VALUES("
                        Sql = Sql & nCodReduz & ",'" & Format(dRef, "mm/dd/yyyy") & "'," & !nroGIA & "," & RetornaNumero(!CFOP) & ","
                        Sql = Sql & Virg2Ponto(!BASECÁLCULO) & "," & Virg2Ponto(!ISENTASNÃOTRIB) & "," & Virg2Ponto(!OUTRAS) & ")"
                        cn.Execute Sql, rdExecDirect
                     End If
proximo:
                   .MoveNext
                Wend
               .Close
            End With
            pBar2.value = 0
            pBar2.Color = vbWhite
           
           .MoveNext
        Wend
       .Close
    End With
PROXIMO2:
Next

PBar.value = 0
PBar.Color = vbWhite
pBar2.value = 0
pBar2.Color = vbWhite

MsgBox "Importação concuída.", vbInformation, "Informação"
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Sub CallPb2(nVal As Long, nTot As Long)
If nVal > 0 Then
    pBar2.Color = &HC0C000
Else
    pBar2.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   pBar2.value = (nVal * 100) / nTot
Else
   pBar2.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Function RetornaCFOP(nCFOP As Integer) As Boolean
Dim Sql As String, Rdoaux As rdoResultset

RetornaCFOP = False

Sql = "SELECT * FROM GIACFOP ORDER BY CFOP1"
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdoaux
    While Not .EOF
        If nCFOP >= !CFOP1 And nCFOP <= !CFOP2 Then
            RetornaCFOP = True
            Exit Function
        End If
       .MoveNext
    Wend
   .Close
End With

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If cnn.State = 1 Then
    cnn.Close
End If
End Sub
