VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProcessoCCusto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processos por Centro de Custos"
   ClientHeight    =   840
   ClientLeft      =   4905
   ClientTop       =   4395
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   6915
   Begin VB.ComboBox cmbReq 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   5550
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      ToolTipText     =   "Imprimir registro"
      Top             =   180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmProcessoCCusto.frx":0000
      PICN            =   "frmProcessoCCusto.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   540
      Width           =   5550
      _ExtentX        =   9790
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
End
Attribute VB_Name = "frmProcessoCCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nNumero As Long, nSeq As Integer, nConta As Integer
Dim ax As String, sNome As String
If cmbReq.ListIndex = -1 Then Exit Sub

nConta = 0
Sql = "SELECT vwFULLPROCESSO.ANO, vwFULLPROCESSO.NUMERO, tramitacao.seq,tramitacao.ccusto, tramitacao.datahora, tramitacao.dataenvio, vwFULLPROCESSO.COMPLEMENTO, "
Sql = Sql & "vwFULLPROCESSO.nomecidadao , vwFULLPROCESSO.DESCRICAO FROM vwFULLPROCESSO INNER JOIN tramitacao ON vwFULLPROCESSO.ANO = tramitacao.ano AND "
Sql = Sql & "vwFULLPROCESSO.NUMERO = tramitacao.numero  Where ccusto = " & cmbReq.ItemData(cmbReq.ListIndex) & " And dataenvio Is Null order by ano ,numero"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    RdoAux.Close
    MsgBox "Este centro de custos não esta com nenhum processo.", vbInformation, "Informação"
Else
    Open sPathBin & "\PROTCC.TXT" For Output As #1
    ax = "PREFEITURA MUNICIPAL DE JABOTICABAL - " & Format(Now, "dd/mm/yyyy") & " " & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00")
    Print #1, ax
    ax = "LISTA DE PROCESSOS QUE SE ENCONTRAM EM PODER DE " & cmbReq.Text
    Print #1, ax

    ax = "Nº PROCESSO     DESDE    DIAS ASSUNTO                             REQUERENTE"
    Print #1, ax
    ax = "====================================================================================================="
    Print #1, ax
    PBar.Color = vbRed
    PBar.value = 0
    
    With RdoAux
        Ocupado
        Do Until .EOF
            CallPb CLng(.AbsolutePosition), CLng(.RowCount)
            DoEvents
            nAno = !Ano
            nNumero = !Numero
            nSeq = !Seq
            Sql = "SELECT * from tramitacao Where tramitacao.Ano = " & nAno & " And tramitacao.Numero = " & nNumero & " And Seq = " & nSeq + 1
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                If Not IsNull(RdoAux2!DATAHORA) Then
                    RdoAux2.Close
                    GoTo PROXIMO
                End If
            End If
            RdoAux2.Close
            nConta = nConta + 1
            If Not IsNull(!nomecidadao) Then
                sNome = !nomecidadao
            Else
                sNome = SubNull(!Descricao)
            End If
            ax = Format(!Numero, "000000") & "-" & RetornaDVProcesso(!Numero) & "/" & !Ano & " " & Format(!DATAHORA, "dd/mm/yyyy") & " " & Format(DateDiff("d", !DATAHORA, Now), "0000") & " " & FillSpace(Left(!Complemento, 35), 35) & " " & FillSpace(Left(sNome, 35), 35)
            Print #1, ax
PROXIMO:
            DoEvents
           .MoveNext
        Loop
       .Close
    End With
    ax = "---------------------"
    Print #1, ax
    ax = "Total de processos: " & nConta
    Print #1, ax
    Close #1
    Liberado
    If nConta > 0 Then
        z = Shell(App.Path & "\NOTEPAD2" & " " & sPathBin & "\PROTCC.TXT", vbNormalFocus)
    End If
End If
PBar.Color = vbWhite
PBar.value = 0

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaCC
End Sub

Private Sub CarregaCC()
Dim Sql As String, RdoAux As rdoResultset
cmbReq.Clear
Sql = "SELECT CODIGO,DESCRICAO FROM CENTROCUSTO where ativo=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbReq.AddItem !Descricao
       cmbReq.ItemData(cmbReq.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

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

