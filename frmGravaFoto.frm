VERSION 5.00
Begin VB.Form frmGravaFoto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grava Fotos "
   ClientHeight    =   4830
   ClientLeft      =   2865
   ClientTop       =   6120
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   7185
   Begin VB.CommandButton Command3 
      Caption         =   "Importar"
      Height          =   345
      Left            =   5730
      TabIndex        =   1
      Top             =   480
      Width           =   1125
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "frmGravaFoto.frx":0000
      Left            =   0
      List            =   "frmGravaFoto.frx":0002
      TabIndex        =   0
      Top             =   90
      Width           =   5385
   End
   Begin Tributacao.XP_ProgressBar pBar 
      Height          =   165
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   291
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
      Color           =   255
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   810
      Top             =   2640
      Width           =   3105
   End
End
Attribute VB_Name = "frmGravaFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Dim lFile As Long, sFullName As String, sNomeArq As String, Sql As String, RdoAux As rdoResultset, nCodReduz As Long, nTot As Long
Dim nDist As Integer, nSetor As Integer, nQuadra As Integer, nLote As Integer, nFace As Integer, nUnidade As Integer, nSubUnidade As Integer, nSeq As Integer
Dim mStream As New ADODB.Stream
Dim rst As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim adoConn As New ADODB.Connection

If Not ConectaBinary Then
    MsgBox "Não foi possivel conectar"
    Exit Sub
End If

adoConn.CursorLocation = adUseClient
adoConn.Open cnBinary.Connect

Set rsEmp = New ADODB.Recordset
 
With rsEmp
  .CursorType = adOpenStatic
  .LockType = adLockOptimistic
  .Open "select * from F001", cnBinary.Connect
End With




pBar.value = 0
nTot = List1.ListCount - 1
For lFile = 0 To nTot
    If lFile Mod 100 = 0 Then
        CallPb lFile, nTot
    End If
    sFullName = List1.List(lFile)
    sNomeArq = GetFileNameFromPath(List1.List(lFile))
    nDist = Left(sNomeArq, 2)
    nSetor = Mid(sNomeArq, 4, 2)
    nQuadra = Mid(sNomeArq, 7, 4)
    nLote = Mid(sNomeArq, 12, 5)
    Sql = "select * from cadimob where distrito=" & nDist & " and setor=" & nSetor & " and quadra=" & nQuadra & " and lote=" & nLote
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        nCodReduz = RdoAux!CODREDUZIDO
    Else
        nCodReduz = 0
    End If
    
    If nCodReduz > 0 Then
        Image1.Picture = LoadPicture(sFullName)
        
        Sql = "Select max(seq) as maximo from F001 where codigo=" & nCodReduz
        Set RdoAux = cnBinary.OpenResultset(Sql, rdOpenKeyset)
        If IsNull(RdoAux!maximo) Then
            nSeq = 0
        Else
            nSeq = RdoAux!maximo + 1
        End If
        
       

        With mStream
            .Type = adTypeBinary
            .Open
            .LoadFromFile sFullName
            rsEmp.AddNew
            rsEmp("codigo") = nCodReduz
            rsEmp("seq") = nSeq
            rsEmp("foto") = .Read
            rsEmp.Update
        End With
        Set mStream = Nothing
        
        
        
    End If
    
Next

MsgBox "fim"
End Sub

Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Private Sub Form_Load()
Dim strStartPath As String
Centraliza Me

strStartPath = "D:\Trabalho\GTI\Fotos\"
ListFolder strStartPath

End Sub

Private Sub ListFolder(sFolderPath As String)
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim subfolder As Folder
    Dim i As Integer
    
    Set FSfolder = FS.GetFolder(sFolderPath)
 
    For Each subfolder In FSfolder.SubFolders
        DoEvents
        i = i + 1
        
        '***
        Dim fso As New FileSystemObject
Dim fld As Folder
Dim fil As File
Set fld = fso.GetFolder(subfolder)
For Each fil In fld.Files
  'Debug.Print fil.Name
  List1.AddItem subfolder & "\" & fil.Name
Next
Set fil = Nothing
Set fld = Nothing
Set fso = Nothing
        
        '*****
        
'        List1.AddItem subfolder
    Next subfolder
    Set FSfolder = Nothing
'    MsgBox "Total sub folders in " & sFolderPath & " : " & i
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    pBar.Color = &HC0C000
Else
    pBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   pBar.value = (nVal * 100) / nTot
Else
   pBar.value = 100
End If

Me.Refresh
DoEvents

End Sub

