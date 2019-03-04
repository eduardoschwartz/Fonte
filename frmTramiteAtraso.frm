VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmTramiteAtraso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trâmites de processo em atraso "
   ClientHeight    =   5385
   ClientLeft      =   11010
   ClientTop       =   5205
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   9960
   Begin vbAcceleratorSGrid6.vbalGrid grdProc 
      Height          =   4695
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   8281
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderDragReorderColumns=   0   'False
      HeaderHotTrack  =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   5040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdFiltrar 
      Default         =   -1  'True
      Height          =   345
      Left            =   8040
      TabIndex        =   2
      ToolTipText     =   "Consulta processos baseados no filtro selecionado"
      Top             =   4950
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Filtrar e imprimir"
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
      MICON           =   "frmTramiteAtraso.frx":0000
      PICN            =   "frmTramiteAtraso.frx":001C
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
Attribute VB_Name = "frmTramiteAtraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFiltrar_Click()
Dim Sql As String, RdoAux As rdoResultset, nAno As Integer, nNumero As Integer, RdoAux2 As rdoResultset, nDespacho As Integer, nPos As Long, nTotal As Long, RdoAux3 As rdoResultset, sCentroCustoDescricao As String, sCentroCustoProximo As String
Exit Sub
nPos = 1
PBar.value = 0
grdProc.Clear
grdProc.Redraw = False
Sql = "select processogti.ano,processogti.NUMERO,processogti.interno,  processogti.DATAENTRADA,processogti.COMPLEMENTO as assunto, processogti.CODCIDADAO, cidadao.nomecidadao, processogti.CENTROCUSTO, centrocusto.DESCRICAO "
Sql = Sql & "From ProcessoGTI inner join centrocusto on processogti.centrocusto = centrocusto.CODIGO left outer join cidadao on processogti.CODCIDADAO=cidadao.codcidadao "
Sql = Sql & "Where ProcessoGTI.Ano = 2018 order by processogti.ano, processogti.NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotal = .RowCount
    Do Until .EOF
        nAno = !Ano
        nNumero = !Numero
        If nPos Mod 50 = 0 Then CallPb nPos, nTotal
        
        sCentroCustoDescricao = ""
        Sql = "select top(1)tramitacaocc.ano,tramitacaocc.numero,ccusto,DESCRICAO from tramitacaocc inner join centrocusto on tramitacaocc.ccusto = centrocusto.CODIGO "
        Sql = Sql & "Where Ano = " & nAno & " And Numero = " & nNumero & " order by seq desc"
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux3.RowCount > 0 Then
            sCentroCustoProximo = RdoAux3!descricao

            Sql = "select top(1) tramitacao.ano,tramitacao.numero,ccusto, centrocusto.DESCRICAO From tramitacao "
            Sql = Sql & "inner join centrocusto on tramitacao.ccusto = centrocusto.CODIGO "
            Sql = Sql & "Where tramitacao.Numero = " & nNumero & " And tramitacao.Ano = " & nAno & " order by tramitacao.seq desc"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                sCentroCustoDescricao = RdoAux!descricao
            End If
        End If
        RdoAux3.Close
        
        
'        Sql = "select top(1) tramitacao.ano,tramitacao.numero,tramitacao.seq,ccusto, centrocusto.DESCRICAO, datahora,despacho,despacho.DESCRICAO as DescDespacho From tramitacao "
'        Sql = Sql & "inner join centrocusto on tramitacao.ccusto = centrocusto.CODIGO inner join despacho on tramitacao.despacho=despacho.CODIGO "
 '       Sql = Sql & "Where tramitacao.Numero = " & nNumero & " And tramitacao.Ano = " & nAno & " order by tramitacao.seq desc"
 '       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        
 '       If RdoAux2.RowCount = 0 Then GoTo Proximo
 '       If Left(RdoAux2!descdespacho, 9) = "ARQUIVADO" Then GoTo Proximo
        
                        
                
        'nDespacho = Val(SubNull(RdoAux2!despacho))
        grdProc.AddRow
        grdProc.CellDetails grdProc.Rows, 1, nAno, DT_CENTER
        grdProc.CellDetails grdProc.Rows, 2, Format(nNumero & RetornaDVProcesso(CLng(nNumero)), "000000-0"), DT_CENTER
        grdProc.CellDetails grdProc.Rows, 3, IIf(!INTERNO, !descricao, !NomeCidadao), DT_LEFT
        grdProc.CellDetails grdProc.Rows, 4, !assunto, DT_LEFT
        grdProc.CellDetails grdProc.Rows, 5, Format(!DATAENTRADA, "dd/mm/yyyy"), DT_CENTER
        'grdProc.CellDetails grdProc.Rows, 6, RdoAux2!descricao, DT_LEFT
        grdProc.CellDetails grdProc.Rows, 6, sCentroCustoDescricao, DT_LEFT
        grdProc.CellDetails grdProc.Rows, 7, "", DT_CENTER
        'grdProc.CellDetails grdProc.Rows, 7, Format(RdoAux2!DATAHORA, "dd/mm/yyyy"), DT_CENTER
        'grdProc.CellDetails grdProc.Rows, 8, RdoAux2!descdespacho, DT_LEFT
        grdProc.CellDetails grdProc.Rows, 8, "", DT_LEFT
        grdProc.CellDetails grdProc.Rows, 9, sCentroCustoProximo, DT_LEFT
        
Proximo:
'        RdoAux2.Close
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

grdProc.Redraw = True
PBar.value = 100
MsgBox "fim"

End Sub

Private Sub Form_Load()
Centraliza Me
GridHeader
End Sub

Private Sub GridHeader()
With grdProc
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    .DefaultRowHeight = 17
    .AddColumn "kAno", "Ano", ecgHdrTextALignCentre, , 40
    .AddColumn "kNum", "Numero", ecgHdrTextALignLeft, , 60
    .AddColumn "kReq", "Requerente", ecgHdrTextALignLeft, , 210
    .AddColumn "kAssu", "Assunto", ecgHdrTextALignLeft, , 200
    .AddColumn "kEnt", "Dt.Entrada", ecgHdrTextALignCentre, , 80
    .AddColumn "kUtra", "Último Trâmite", ecgHdrTextALignLeft, , 200
    .AddColumn "kDul", "Dt.Último", ecgHdrTextALignCentre, , 80
    .AddColumn "kDsu", "Despacho último", ecgHdrTextALignLeft, , 150
    .AddColumn "kPtra", "Próximo Trâmite", ecgHdrTextALignLeft, , 200
End With

End Sub

Private Sub CallPb(nPos As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents
If nTotal = 0 Then Exit Sub
If ((nPos * 100) / nTotal) <= 100 Then
   PBar.value = (nPos * 100) / nTotal
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub


