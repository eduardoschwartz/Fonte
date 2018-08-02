VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmDocNF 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Fiscais por Documento"
   ClientHeight    =   3075
   ClientLeft      =   3810
   ClientTop       =   5640
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   9855
   Begin VB.TextBox txtNumDoc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   0
      Top             =   90
      Width           =   1395
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Default         =   -1  'True
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Consulta notas fiscais do documento"
      Top             =   90
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmDocNF.frx":0000
      PICN            =   "frmDocNF.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdMain 
      Height          =   2535
      Left            =   45
      TabIndex        =   3
      Top             =   495
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4471
      NoHorizontalGridLines=   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   16777215
      HighlightBackColor=   128
      HighlightForeColor=   16777215
      GroupRowForeColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DrawFocusRectangle=   0   'False
      GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Documento:"
      Height          =   225
      Left            =   135
      TabIndex        =   2
      Top             =   150
      Width           =   1440
   End
End
Attribute VB_Name = "frmDocNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NOTAS
    IdentificaPrestador As String
    TipoPrestador As Integer
    TipoNota As Integer
    NumeroNota As Long
    Serie As String
    DataEmissao As String
    MesRef As Integer
    AnoRef As Integer
    StatusNota As Integer
    DataCancel As String
    Natureza As String
    ValorTotal As Double
    ValorServico As Double
    ValorImposto As Double
    Recolhimento As Integer
    Atividade As Integer
    Aliquota As Double
    RazaoPrestador As String
    CidadePrestador As String
    UFPrestador As String
    LocalPrestador As String
    IdentificaTomador As String
    TipoTomador As String
    RazaoTomador As String
    CidadeTomador As String
    UFTomador As String
    LocalTomador As String
End Type
Dim aNota() As NOTAS

Public Sub cmdConsultar_Click()
CarregaISS
CarregaNF
End Sub

Private Sub Form_Load()
GridHeader
ReDim aNota(0)
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)

Tweak txtNumDoc, KeyAscii, IntegerPositive
End Sub

Private Sub CarregaNF()
Dim x As Integer, Sql As String, RdoAux As rdoResultset


Ocupado
grdMain.Redraw = False
grdMain.Clear
grdMain.Redraw = True
grdMain.Redraw = False

For x = 1 To UBound(aNota)
    With aNota(x)
'        If .AnoRef = Val(cmbAnoISS.text) And .MesRef = Val(lblMesNF.Caption) Then
            grdMain.AddRow
            grdMain.CellDetails grdMain.Rows, 1, .NumeroNota, DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 2, .Serie, DT_CENTER
            If .TipoNota = 1 Then
                grdMain.CellDetails grdMain.Rows, 3, "Emitida", DT_LEFT
            Else
                grdMain.CellDetails grdMain.Rows, 3, "Recebida", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 4, .DataEmissao, DT_CENTER
            If .StatusNota = 0 Then
                grdMain.CellDetails grdMain.Rows, 5, "Recebida", DT_LEFT
            ElseIf .StatusNota = 1 Then
                grdMain.CellDetails grdMain.Rows, 5, "Normal", DT_LEFT
            ElseIf .StatusNota = 2 Then
                grdMain.CellDetails grdMain.Rows, 5, "Cancelada", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 6, .DataCancel, DT_CENTER
            If .Natureza = 1 Then
                grdMain.CellDetails grdMain.Rows, 7, "Serviço", DT_CENTER
            Else
                grdMain.CellDetails grdMain.Rows, 7, "Mista", DT_CENTER
            End If
            grdMain.CellDetails grdMain.Rows, 8, FormatNumber(.ValorTotal, 2), DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 9, FormatNumber(.ValorServico, 2), DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 10, FormatNumber(.ValorImposto, 2), DT_RIGHT
            If .TipoNota = 1 And .Recolhimento = 0 Then
                grdMain.CellDetails grdMain.Rows, 11, "Isento", DT_LEFT
            ElseIf .TipoNota = 1 And .Recolhimento = 1 Then
                grdMain.CellDetails grdMain.Rows, 11, "Retido", DT_LEFT
            ElseIf .TipoNota = 1 And .Recolhimento = 2 Then
                grdMain.CellDetails grdMain.Rows, 11, "A Recolher", DT_LEFT
            ElseIf .TipoNota = 1 And .Recolhimento = 3 Then
                grdMain.CellDetails grdMain.Rows, 11, "Simples", DT_LEFT
            ElseIf .TipoNota = 2 And .Recolhimento = 1 Then
                grdMain.CellDetails grdMain.Rows, 11, "Disp.Ret.", DT_LEFT
            ElseIf .TipoNota = 2 And .Recolhimento = 2 Then
                grdMain.CellDetails grdMain.Rows, 11, "Ret.Sub.Trib.", DT_LEFT
            ElseIf .TipoNota = 2 And .Recolhimento = 3 Then
                grdMain.CellDetails grdMain.Rows, 11, "Ret.Res.Trib.", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 12, .Atividade, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 13, FormatNumber(.Aliquota, 2) & "%", DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 14, .IdentificaPrestador, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 15, .RazaoPrestador, DT_LEFT
            If .TipoTomador = 0 Then
                grdMain.CellDetails grdMain.Rows, 16, Format(.IdentificaTomador, "0#\.###\.###/####-##"), DT_LEFT
                grdMain.CellDetails grdMain.Rows, 17, "CNPJ", DT_LEFT
            ElseIf .TipoTomador = 1 Then
                grdMain.CellDetails grdMain.Rows, 16, Format(.IdentificaTomador, "00#\.###\.###-##"), DT_LEFT
                grdMain.CellDetails grdMain.Rows, 17, "CPF", DT_LEFT
            ElseIf .TipoTomador = 2 Then
                grdMain.CellDetails grdMain.Rows, 16, .IdentificaTomador, DT_LEFT
                grdMain.CellDetails grdMain.Rows, 17, "IM", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 18, .RazaoTomador, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 19, .CidadeTomador & "/" & .UFTomador, DT_LEFT
 '       End If
    End With
Next
Liberado
grdMain.Redraw = True

End Sub

Private Sub GridHeader()

With grdMain
    .HeaderFlat = True
    .HeaderHeight = 18
    .DefaultRowHeight = 17
    .GridFillLineColor = vbWhite
    .RowMode = True
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
        
    .AddColumn "NumNota", "N° Nota", ecgHdrTextALignRight, , 50
    .AddColumn "Serie", "Série", ecgHdrTextALignCentre, , 40
    .AddColumn "TipoNota", "Tipo", ecgHdrTextALignLeft, , 50
    .AddColumn "DtEmissao", "Dt.Emissão", ecgHdrTextALignCentre, , 70
    .AddColumn "Situação", "Situação", ecgHdrTextALignLeft, , 60
    .AddColumn "DtCancel", "Dt.Cancel.", ecgHdrTextALignCentre, , 70
    .AddColumn "Natureza", "Natureza", ecgHdrTextALignLeft, , 60
    .AddColumn "VlTotal", "Vl.Total", ecgHdrTextALignRight, , 70
    .AddColumn "VlServico", "Vl.Serviço", ecgHdrTextALignRight, , 70
    .AddColumn "VlImposto", "Vl.Imposto", ecgHdrTextALignRight, , 70
    .AddColumn "Recolh", "Recolhim.", ecgHdrTextALignLeft, , 70
    .AddColumn "Atividade", "Atividade", ecgHdrTextALignLeft, , 60
    .AddColumn "Aliq", "Aliq", ecgHdrTextALignRight, , 40
    .AddColumn "IdPrestador", "Id.Prestador", ecgHdrTextALignLeft, , 110
    .AddColumn "RazaoPrestador", "Razão Prestador", ecgHdrTextALignLeft, , 130
    .AddColumn "IdTomador", "Id.Tomador", ecgHdrTextALignLeft, , 110
    .AddColumn "TipoTomador", "Tipo", ecgHdrTextALignLeft, , 40
    .AddColumn "RazaoTomador", "Razão Tomador", ecgHdrTextALignLeft, , 130
    .AddColumn "CidadeTomador", "Cidade/UF", ecgHdrTextALignLeft, , 130
End With

End Sub

Private Sub CarregaISS()
Dim Sql As String, RdoAux As rdoResultset, nPos As Integer
Dim x As Integer, bAchou As Boolean, nTotalE As Double, nTotalR As Double
If Val(txtNumDoc.Text) = 0 Then Exit Sub
Ocupado
On Error Resume Next
ReDim aNota(0): nPos = 1: ReDim aISS(0): nTotalE = 0: nTotalR = 0
Sql = "SELECT * FROM NFISSELETRO2 WHERE NUMDOC=" & Val(txtNumDoc.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aNota(nPos)
        aNota(nPos).NumeroNota = !NumeroNota
        aNota(nPos).Serie = !Serie
        aNota(nPos).TipoNota = !TipoNota
        aNota(nPos).AnoRef = !AnoRef
        aNota(nPos).MesRef = !MesRef
        aNota(nPos).DataEmissao = Format(!DataEmissao, "dd/mm/yyyy")
        aNota(nPos).StatusNota = !StatusNota
        aNota(nPos).Natureza = !Natureza
        aNota(nPos).DataCancel = Format(!DataCancel, "dd/mm/yyyy")
        aNota(nPos).ValorTotal = !ValorTotal
        aNota(nPos).ValorServico = !ValorServico
        aNota(nPos).ValorImposto = !ValorImposto
        aNota(nPos).Recolhimento = !Recolhimento
        aNota(nPos).Atividade = !Atividade
        aNota(nPos).Aliquota = !Aliquota
        aNota(nPos).IdentificaPrestador = !IdentificaPrestador
        aNota(nPos).RazaoPrestador = !RazaoPrestador
        aNota(nPos).IdentificaTomador = !IdentificaTomador
        aNota(nPos).TipoTomador = !TipoTomador
        aNota(nPos).RazaoTomador = !RazaoTomador
        aNota(nPos).CidadeTomador = !CidadeTomador
        aNota(nPos).UFTomador = !UFTomador
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

End Sub

